import { decodeBase64 } from '../../../common/binary/base64';

import { parseXlsx } from '../parser';
import { resolveXlsxResourceLimits } from '../internal/resource-limits';
import { canonicalXlsxJson } from './canonical-json';
import { createXlsxCapabilityManifest } from './capability';
import { canonicalXlsxSha256 } from './digest';
import { XlsxWriteError } from './errors';
import {
  inspectXlsxPackageGraph,
  type XlsxPackageGraph,
} from './internal/package-graph';
import { createXlsxRoundTripDocument } from './keys';
import { normalizeXlsxRoundTripSource } from './source';
import type {
  ResolvedXlsxWriteLimits,
  XlsxRoundTripSnapshot,
  XlsxWriteOptions,
} from './types';
import { validateXlsxSnapshotShape } from './validate-snapshot';

export interface VerifiedXlsxSnapshot {
  bytes: Uint8Array;
  graph: XlsxPackageGraph;
  snapshot: XlsxRoundTripSnapshot;
}

function integrityFailure(message: string): never {
  throw new XlsxWriteError('snapshot-integrity-failed', message);
}

function sourceFailure(message: string): never {
  throw new XlsxWriteError('source-package-mismatch', message);
}

export async function verifyXlsxRoundTripSnapshot(
  value: unknown,
  options: XlsxWriteOptions,
  writeLimits: ResolvedXlsxWriteLimits,
): Promise<VerifiedXlsxSnapshot> {
  const snapshot = validateXlsxSnapshotShape(value, writeLimits);
  const readerLimits = resolveXlsxResourceLimits(options.readerLimits);
  const decoded = decodeBase64(snapshot.source.packageBase64);
  const normalized = await normalizeXlsxRoundTripSource(
    decoded,
    readerLimits,
    writeLimits,
  );
  if (normalized.sha256 !== snapshot.source.sha256) {
    sourceFailure('XLSX source bytes do not match their snapshot identity');
  }

  let graph: XlsxPackageGraph;
  let parsed: Awaited<ReturnType<typeof parseXlsx>>;
  try {
    graph = await inspectXlsxPackageGraph(normalized.bytes, readerLimits);
    parsed = await parseXlsx(normalized.bytes, {
      errorMode: 'strict',
      imageMode: 'none',
      limits: options.readerLimits ?? {},
      pivotCacheMode: 'metadata',
    });
  } catch {
    sourceFailure('XLSX source package failed strict verification');
  }
  if (
    graph.conformance !== snapshot.source.conformance ||
    graph.manifestHash !== snapshot.sourceManifestHash
  ) {
    sourceFailure('XLSX source package graph does not match the snapshot');
  }
  const preservation = {
    containsActiveContent: graph.containsActiveContent,
    containsDigitalSignatures: graph.containsDigitalSignatures,
    containsExternalRelationships: graph.containsExternalRelationships,
    containsOpaqueContent: graph.containsOpaqueContent,
    securityMode: snapshot.preservation.securityMode,
  };
  if (
    canonicalXlsxJson(preservation) !== canonicalXlsxJson(snapshot.preservation)
  ) {
    sourceFailure('XLSX preservation inventory does not match the source');
  }
  if (
    graph.containsActiveContent &&
    snapshot.preservation.securityMode === 'reject-active'
  ) {
    throw new XlsxWriteError(
      'preservation-conflict',
      'XLSX source contains active or embedded content',
      { featureClass: 'active-content' },
    );
  }
  const profile = createXlsxCapabilityManifest();
  const document = await createXlsxRoundTripDocument(
    parsed,
    normalized.sha256,
    profile.id,
  );
  const baseDocumentHash = await canonicalXlsxSha256(document);
  if (
    baseDocumentHash !== snapshot.baseDocumentHash ||
    baseDocumentHash !== snapshot.stateHash ||
    canonicalXlsxJson(document) !== canonicalXlsxJson(snapshot.document)
  ) {
    integrityFailure('XLSX semantic preview does not match its source');
  }
  return { bytes: normalized.bytes, graph, snapshot };
}
