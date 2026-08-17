import { readZipEntryBytes } from '../../../common/archive/read-entry';
import {
  assertXmlComplexity,
  decodeXmlBytes,
} from '../../../common/xml/validate';
import type { ResolvedPptxResourceLimits } from '../internal/resource-limits';
import { PptxWriteError } from '../write-error';
import { unsupportedPptxEdit } from './patch-error';
import JSZip from 'jszip';

function byteEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((value, index) => value === right[index]);
}

export async function readPptxPartPayloads(
  archive: JSZip,
  limits: ResolvedPptxResourceLimits,
): Promise<Map<string, Uint8Array>> {
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (file.dir) continue;
    result.set(file.name, await readZipEntryBytes(file, limits.maxPartBytes));
  }
  return result;
}

export function assertSafeEditablePptxPackage(archive: JSZip): void {
  for (const name of Object.keys(archive.files)) {
    const normalized = name.toLowerCase();
    if (
      normalized.startsWith('_xmlsignatures/') ||
      normalized.endsWith('/vbaproject.bin')
    ) {
      unsupportedPptxEdit(
        'PowerPoint text edit does not modify signed or macro-enabled packages',
      );
    }
  }
}

export function decodeEditablePptxXml(
  bytes: Uint8Array,
  limits: ResolvedPptxResourceLimits,
): string {
  if (hasPptxUtf16Bom(bytes)) {
    unsupportedPptxEdit('PowerPoint text edit requires UTF-8 slide XML');
  }
  const xml = decodeXmlBytes(bytes);
  assertXmlComplexity(xml, {
    maxDepth: limits.maxXmlDepth,
    maxNodes: limits.maxXmlNodes,
  });
  if (/encoding\s*=\s*["']utf-16["']/i.test(xml)) {
    unsupportedPptxEdit('PowerPoint text edit requires UTF-8 slide XML');
  }
  return xml;
}

export function hasPptxUtf16Bom(bytes: Uint8Array): boolean {
  return (
    (bytes[0] === 0xff && bytes[1] === 0xfe) ||
    (bytes[0] === 0xfe && bytes[1] === 0xff)
  );
}

export function verifyPptxPatchedPayloads(
  sourcePayloads: ReadonlyMap<string, Uint8Array>,
  outputPayloads: ReadonlyMap<string, Uint8Array>,
  patchedParts: ReadonlySet<string>,
): void {
  if (
    outputPayloads.size !== sourcePayloads.size ||
    [...sourcePayloads.keys()].some((name) => !outputPayloads.has(name))
  ) {
    throw new PptxWriteError(
      'verification-failed',
      'PowerPoint text edit changed the package part inventory',
    );
  }
  for (const [name, source] of sourcePayloads) {
    const result = outputPayloads.get(name) as Uint8Array;
    const equal = byteEqual(source, result);
    if (patchedParts.has(name) ? equal : !equal) {
      throw new PptxWriteError(
        'verification-failed',
        patchedParts.has(name)
          ? `PowerPoint text edit did not change dirty part ${name}`
          : `PowerPoint text edit changed untouched part ${name}`,
      );
    }
  }
}

export function generatePptxPatchedArchive(
  archive: JSZip,
): Promise<Uint8Array> {
  return archive.generateAsync({
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
    platform: 'DOS',
    streamFiles: false,
    type: 'uint8array',
  });
}
