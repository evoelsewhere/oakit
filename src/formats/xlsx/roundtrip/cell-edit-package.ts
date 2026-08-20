import { readZipEntryBytes } from '../../../common/archive/read-entry';
import JSZip from 'jszip';

import { XlsxPartReader } from '../internal/part-reader';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import { discoverXlsxWorkbook } from '../internal/workbook-discovery';
import { parseXlsxWorkbookManifest } from '../internal/workbook-manifest';
import { loadXlsxStyles, type XlsxStyleTable } from '../internal/styles';
import { XlsxWriteError } from './errors';
import {
  inspectXlsxPackageGraph,
  type XlsxPackageGraph,
} from './internal/package-graph';
import {
  assertXlsxCellEditFormulaClosure,
  assertXlsxCellEditStyleClosure,
  assertXlsxInternalHyperlinkEditClosure,
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
  appendXlsxStylesPart,
  xlsxAppendedStyleRecordCount,
} from './style-append';
import {
  patchXlsxInternalHyperlinks,
  type XlsxInternalHyperlinkPatch,
} from './hyperlink-patch';
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
  appendedStyleXfs: ReadonlyMap<number, number>,
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
    if (impact.kind === 'set-hyperlink') continue;
    let sheet = bySheet.get(impact.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(impact.sheetKey, sheet);
    }
    const cell = xlsxPlannedCell(plan.document, impact.sheetKey, impact.cell);
    const styleCell = styleCells.has(`${impact.sheetKey}\u0000${impact.cell}`);
    const xmlStyleIndex = styleCell
      ? cell.style !== undefined && appendedStyleXfs.has(cell.style)
        ? appendedStyleXfs.get(cell.style)
        : styles.cellXfs.findIndex(
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

function finalHyperlinkPatches(
  plan: XlsxCellOperationPlan,
): Map<string, XlsxInternalHyperlinkPatch[]> {
  const bySheet = new Map<string, Map<string, XlsxInternalHyperlinkPatch>>();
  for (const operation of plan.operations) {
    if (operation.kind !== 'set-hyperlink') continue;
    let sheet = bySheet.get(operation.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(operation.sheetKey, sheet);
    }
    sheet.set(operation.cell, {
      cell: operation.cell,
      operationId: operation.operationId,
      target: operation.target as XlsxInternalHyperlinkPatch['target'],
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
  assertXlsxInternalHyperlinkEditClosure(baseDocument, plan);
  const context = await packageContext(sourceBytes, readerLimits);
  const appendedStyles = plan.document.styles.slice(baseDocument.styles.length);
  const outputStyleRecords =
    context.styles.recordCount + xlsxAppendedStyleRecordCount(appendedStyles);
  if (outputStyleRecords > readerLimits.maxStyles) {
    throw new XlsxWriteError(
      'resource-limit-exceeded',
      'XLSX appended style records exceed the reader limit',
      {
        actual: outputStyleRecords,
        limit: readerLimits.maxStyles,
        limitName: 'maxStyles',
      },
    );
  }
  if (appendedStyles.length !== 0 && context.styles.part === null) {
    throw new XlsxWriteError(
      'preservation-conflict',
      'XLSX cannot append styles without an existing styles part',
      { featureClass: 'missing-styles-part' },
    );
  }
  const dirtyPartCount =
    new Set(plan.impacts.map((impact) => impact.sheetKey)).size +
    (appendedStyles.length === 0 ? 0 : 1);
  if (dirtyPartCount > writeLimits.maxDirtyParts) {
    writeLimitFailure(
      'maxDirtyParts',
      dirtyPartCount,
      writeLimits.maxDirtyParts,
    );
  }
  if (dirtyPartCount > writeLimits.maxPatchedParts) {
    writeLimitFailure(
      'maxPatchedParts',
      dirtyPartCount,
      writeLimits.maxPatchedParts,
    );
  }
  const dependencyEdges = plan.impacts.length + appendedStyles.length;
  if (dependencyEdges > writeLimits.maxDependencyEdges) {
    writeLimitFailure(
      'maxDependencyEdges',
      dependencyEdges,
      writeLimits.maxDependencyEdges,
    );
  }

  const archive = await JSZip.loadAsync(sourceBytes);
  const appendedStyleXfs = new Map<number, number>();
  const dirtyParts = new Set<string>();
  let generatedXmlBytes = 0;
  let patchBytes = 0;
  let patchCount = 0;
  if (appendedStyles.length !== 0) {
    const part = context.styles.part!;
    const entry = archive.file(part)!;
    const source = await readZipEntryBytes(entry, readerLimits.maxPartBytes);
    const appended = appendXlsxStylesPart(
      source,
      appendedStyles,
      writeLimits,
      part,
    );
    generatedXmlBytes += appended.data.byteLength;
    patchBytes += appended.patchBytes;
    patchCount += appended.patchCount;
    archive.file(part, appended.data, { date: entry.date });
    dirtyParts.add(part);
    for (const [offset, xmlStyleIndex] of appended.cellXfIndexes.entries()) {
      appendedStyleXfs.set(baseDocument.styles.length + offset, xmlStyleIndex);
    }
  }
  const patches = finalPatches(plan, context.styles, appendedStyleXfs);
  const hyperlinkPatches = finalHyperlinkPatches(plan);
  const sheetKeys = new Set([...patches.keys(), ...hyperlinkPatches.keys()]);
  for (const sheetKey of sheetKeys) {
    const sheetPatches = patches.get(sheetKey) ?? [];
    const sheet = baseDocument.sheets.find(
      (candidate) => candidate.key === sheetKey,
    )!;
    const part = context.sheetParts[sheet.index]!;
    const entry = archive.file(part)!;
    const source = await readZipEntryBytes(entry, readerLimits.maxPartBytes);
    const cellPatched = patchXlsxWorksheetPartWithReport(
      source,
      sheetPatches,
      writeLimits,
      part,
    );
    const hyperlinkPatched = patchXlsxInternalHyperlinks(
      cellPatched.data,
      hyperlinkPatches.get(sheetKey) ?? [],
      writeLimits,
      part,
    );
    generatedXmlBytes += hyperlinkPatched.data.byteLength;
    if (generatedXmlBytes > writeLimits.maxGeneratedXmlBytes) {
      writeLimitFailure(
        'maxGeneratedXmlBytes',
        generatedXmlBytes,
        writeLimits.maxGeneratedXmlBytes,
        part,
      );
    }
    patchBytes += cellPatched.patchBytes + hyperlinkPatched.patchBytes;
    if (patchBytes > writeLimits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        writeLimits.maxPatchBytes,
        part,
      );
    }
    patchCount += cellPatched.patchCount + hyperlinkPatched.patchCount;
    if (patchCount > writeLimits.maxPatchCount) {
      writeLimitFailure(
        'maxPatchCount',
        patchCount,
        writeLimits.maxPatchCount,
        part,
      );
    }
    archive.file(part, hyperlinkPatched.data, { date: entry.date });
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
