import { readZipEntryBytes } from '../../../common/archive/read-entry';
import JSZip from 'jszip';

import { XlsxPartReader } from '../internal/part-reader';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import { discoverXlsxWorkbook } from '../internal/workbook-discovery';
import { parseXlsxWorkbookManifest } from '../internal/workbook-manifest';
import { loadXlsxStyles, type XlsxStyleTable } from '../internal/styles';
import {
  inspectXlsxPackageGraph,
  type XlsxPackageGraph,
} from './internal/package-graph';
import {
  assertXlsxCellEditFormulaClosure,
  assertXlsxCellEditStyleClosure,
  assertXlsxSafeCellEditSource,
  xlsxPlannedCell,
} from './cell-edit-policy';
import {
  generateBoundedXlsxZip,
  verifyXlsxCellEditR1Parts,
} from './cell-edit-verification';
import type { XlsxCellOperationPlan } from './operation-planner';
import type {
  ResolvedXlsxWriteLimits,
  XlsxPartFidelity,
  XlsxRoundTripDocument,
  XlsxWriteOptions,
} from './types';
import {
  patchXlsxWorksheetPartWithReport,
  type XlsxWorksheetCellPatch,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

export interface XlsxCellEditPackage {
  data: Uint8Array;
  graph: XlsxPackageGraph;
  parts: XlsxPartFidelity[];
}

async function packageContext(
  bytes: Uint8Array,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<{ sheetParts: string[]; styles: XlsxStyleTable }> {
  const archive = await JSZip.loadAsync(bytes);
  const reader = new XlsxPartReader(archive, [], readerLimits);
  const discovery = await discoverXlsxWorkbook(reader, readerLimits);
  const manifest = await parseXlsxWorkbookManifest(
    discovery,
    reader,
    readerLimits,
  );
  return {
    sheetParts: manifest.sheetParts,
    styles: await loadXlsxStyles(discovery, reader, readerLimits),
  };
}

function finalPatches(
  plan: XlsxCellOperationPlan,
  styles: XlsxStyleTable,
): Map<string, XlsxWorksheetCellPatch[]> {
  const bySheet = new Map<string, Map<string, XlsxWorksheetCellPatch>>();
  const styleCells = new Set(
    plan.impacts
      .filter((impact) => impact.kind === 'set-cell-style')
      .map((impact) => `${impact.sheetKey}\u0000${impact.cell}`),
  );
  const contentCells = new Set(
    plan.impacts
      .filter((impact) => impact.kind !== 'set-cell-style')
      .map((impact) => `${impact.sheetKey}\u0000${impact.cell}`),
  );
  for (const impact of plan.impacts) {
    let sheet = bySheet.get(impact.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(impact.sheetKey, sheet);
    }
    const cell = xlsxPlannedCell(plan.document, impact.sheetKey, impact.cell);
    const styleCell = styleCells.has(`${impact.sheetKey}\u0000${impact.cell}`);
    const xmlStyleIndex = styleCell
      ? styles.cellXfs.findIndex(
          (candidate) => candidate.normalizedStyle === cell.style,
        )
      : undefined;
    sheet.set(impact.cell, {
      cell,
      contentChanged: contentCells.has(
        `${impact.sheetKey}\u0000${impact.cell}`,
      ),
      operationId: impact.operationId,
      ...(xmlStyleIndex === undefined ? {} : { xmlStyleIndex }),
    });
  }
  return new Map(
    [...bySheet].map(([sheetKey, patches]) => [
      sheetKey,
      [...patches.values()],
    ]),
  );
}

export async function writeXlsxCellEditPackage(
  sourceBytes: Uint8Array,
  sourceGraph: XlsxPackageGraph,
  baseDocument: XlsxRoundTripDocument,
  plan: XlsxCellOperationPlan,
  options: XlsxWriteOptions,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<XlsxCellEditPackage> {
  assertXlsxSafeCellEditSource(sourceGraph, options);
  assertXlsxCellEditFormulaClosure(baseDocument, plan);
  assertXlsxCellEditStyleClosure(baseDocument, plan);
  const context = await packageContext(sourceBytes, readerLimits);
  const patches = finalPatches(plan, context.styles);
  if (patches.size > writeLimits.maxDirtyParts) {
    writeLimitFailure('maxDirtyParts', patches.size, writeLimits.maxDirtyParts);
  }
  if (patches.size > writeLimits.maxPatchedParts) {
    writeLimitFailure(
      'maxPatchedParts',
      patches.size,
      writeLimits.maxPatchedParts,
    );
  }
  if (plan.impacts.length > writeLimits.maxDependencyEdges) {
    writeLimitFailure(
      'maxDependencyEdges',
      plan.impacts.length,
      writeLimits.maxDependencyEdges,
    );
  }
  const patchCount = [...patches.values()].reduce(
    (total, sheetPatches) => total + sheetPatches.length,
    0,
  );
  if (patchCount > writeLimits.maxPatchCount) {
    writeLimitFailure('maxPatchCount', patchCount, writeLimits.maxPatchCount);
  }

  const archive = await JSZip.loadAsync(sourceBytes);
  const dirtyParts = new Set<string>();
  let generatedXmlBytes = 0;
  let patchBytes = 0;
  for (const [sheetKey, sheetPatches] of patches) {
    const sheet = baseDocument.sheets.find(
      (candidate) => candidate.key === sheetKey,
    )!;
    const part = context.sheetParts[sheet.index]!;
    const entry = archive.file(part)!;
    const source = await readZipEntryBytes(entry, readerLimits.maxPartBytes);
    const patched = patchXlsxWorksheetPartWithReport(
      source,
      sheetPatches,
      writeLimits,
      part,
    );
    generatedXmlBytes += patched.data.byteLength;
    if (generatedXmlBytes > writeLimits.maxGeneratedXmlBytes) {
      writeLimitFailure(
        'maxGeneratedXmlBytes',
        generatedXmlBytes,
        writeLimits.maxGeneratedXmlBytes,
        part,
      );
    }
    patchBytes += patched.patchBytes;
    if (patchBytes > writeLimits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        writeLimits.maxPatchBytes,
        part,
      );
    }
    archive.file(part, patched.data, { date: entry.date });
    dirtyParts.add(part);
  }
  const data = await generateBoundedXlsxZip(
    archive,
    writeLimits.maxOutputBytes,
  );
  const graph = await inspectXlsxPackageGraph(data, readerLimits);
  const fidelityParts = verifyXlsxCellEditR1Parts(
    sourceGraph,
    graph,
    dirtyParts,
  );
  return { data, graph, parts: fidelityParts };
}
