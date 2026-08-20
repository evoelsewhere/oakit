import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from '../internal/cell-reference';
import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import { transformXlsxStructuralRange } from './structural-reference';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlAttributeSpan,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

export interface XlsxWorksheetStructurePatch {
  count: number;
  index: number;
  kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
  operationId: string;
}

export interface XlsxWorksheetStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

function failure(
  message: string,
  part: string,
  request?: XlsxWorksheetStructurePatch,
  featureClass = 'worksheet-structure-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
    ...(request === undefined
      ? {}
      : {
          operationId: request.operationId,
          range: `${request.index}:${request.index + request.count - 1}`,
        }),
    part,
  });
}

function attribute(
  token: XlsxXmlTagToken,
  name: string,
): XlsxXmlAttributeSpan | undefined {
  return token.attributes.find((candidate) => candidate.name === name);
}

function attributePatch(span: XlsxXmlAttributeSpan, value: string): TextPatch {
  return {
    end: span.end,
    replacement: ` ${span.name}="${value}"`,
    start: span.start,
  };
}

function directRows(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  part: string,
): { rows: XlsxXmlTagToken[]; sheetData: XlsxXmlTagToken } {
  const prefix = root.name.slice(0, -'worksheet'.length);
  const sheetData = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}sheetData`,
  );
  if (!sheetData) {
    failure('XLSX worksheet sheetData cannot patch structure', part);
  }
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(sheetData));
  return {
    rows: tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === sheetData.depth + 1 &&
        token.name === `${prefix}row` &&
        token.start >= sheetData.end &&
        token.end <= close.start,
    ),
    sheetData,
  };
}

function shiftedIndex(
  value: number,
  request: XlsxWorksheetStructurePatch,
): number | null {
  const end = request.index + request.count - 1;
  if (request.kind.startsWith('insert-')) {
    if (value < request.index) return value;
    return value + request.count;
  }
  if (value < request.index) return value;
  if (value <= end) return null;
  return value - request.count;
}

function layoutPatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  prefix: string,
  request: XlsxWorksheetStructurePatch,
  part: string,
): TextPatch[] {
  const patches: TextPatch[] = [];
  const dimension = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}dimension`,
  );
  if (dimension) {
    const reference = attribute(dimension, 'ref');
    const range = parseXlsxRangeReference(reference?.value);
    if (!reference || !range) {
      failure('XLSX structural dimension reference is invalid', part, request);
    }
    const transformed = transformXlsxStructuralRange(range, request);
    if (transformed === null) {
      const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(dimension));
      patches.push({ end: close.end, replacement: '', start: dimension.start });
    } else if (transformed.reference !== range.reference) {
      patches.push(attributePatch(reference, transformed.reference));
    }
  }
  const mergeCells = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}mergeCells`,
  );
  if (!mergeCells) return patches;
  const mergeClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(mergeCells));
  const entries = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === mergeCells.depth + 1 &&
      token.name === `${prefix}mergeCell` &&
      token.start >= mergeCells.end &&
      token.end <= mergeClose.start,
  );
  const transformedEntries = entries.map((entry) => {
    const reference = attribute(entry, 'ref');
    const range = parseXlsxRangeReference(reference?.value);
    if (!reference || !range) {
      failure('XLSX structural merged range is invalid', part, request);
    }
    return {
      entry,
      range,
      reference,
      transformed: transformXlsxStructuralRange(range, request),
    };
  });
  const remaining = transformedEntries.filter(
    (entry) => entry.transformed !== null,
  );
  if (remaining.length === 0) {
    patches.push({
      end: mergeClose.end,
      replacement: '',
      start: mergeCells.start,
    });
    return patches;
  }
  for (const item of transformedEntries) {
    if (item.transformed === null) {
      const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(item.entry));
      patches.push({
        end: close.end,
        replacement: '',
        start: item.entry.start,
      });
    } else if (item.transformed.reference !== item.range.reference) {
      patches.push(attributePatch(item.reference, item.transformed.reference));
    }
  }
  const count = attribute(mergeCells, 'count');
  if (count && count.value !== String(remaining.length)) {
    patches.push(attributePatch(count, String(remaining.length)));
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'worksheet',
  );
  if (!root)
    failure('XLSX worksheet root cannot patch structure', part, request);
  const { rows } = directRows(tokens, root, part);
  const prefix = root.name.slice(0, -'worksheet'.length);
  const patches = layoutPatches(tokens, root, prefix, request, part);
  for (const row of rows) {
    const rowReference = attribute(row, 'r');
    const rowIndex = Number(rowReference?.value);
    if (!rowReference || !Number.isSafeInteger(rowIndex)) {
      failure('XLSX structural target row reference is invalid', part, request);
    }
    const rowClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(row));
    if (request.kind.endsWith('-rows')) {
      const shifted = shiftedIndex(rowIndex, request);
      if (shifted === null) {
        patches.push({ end: rowClose.end, replacement: '', start: row.start });
        continue;
      }
      if (shifted !== rowIndex) {
        patches.push(attributePatch(rowReference, String(shifted)));
      }
    }
    if (request.kind.endsWith('-columns')) {
      const spans = attribute(row, 'spans');
      if (spans)
        patches.push({ end: spans.end, replacement: '', start: spans.start });
    }
    const cells = tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === row.depth + 1 &&
        token.name === `${prefix}c` &&
        token.start >= row.end &&
        token.end <= rowClose.start,
    );
    for (const cell of cells) {
      const reference = attribute(cell, 'r');
      const parsed = parseXlsxCellReference(reference?.value);
      if (!reference || !parsed) {
        failure(
          'XLSX structural target cell reference is invalid',
          part,
          request,
        );
      }
      if (request.kind.endsWith('-rows')) {
        const shiftedRow = shiftedIndex(parsed.row, request)!;
        if (shiftedRow !== parsed.row) {
          patches.push(
            attributePatch(
              reference,
              `${xlsxColumnName(parsed.column)!}${shiftedRow}`,
            ),
          );
        }
        continue;
      }
      const shiftedColumn = shiftedIndex(parsed.column, request);
      if (shiftedColumn === null) {
        const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(cell));
        patches.push({ end: close.end, replacement: '', start: cell.start });
      } else if (shiftedColumn !== parsed.column) {
        patches.push(
          attributePatch(
            reference,
            `${xlsxColumnName(shiftedColumn)!}${parsed.row}`,
          ),
        );
      }
    }
  }
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
  }
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of patches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const data = encodeXlsxXml({ ...decoded, text: output });
  if (data.byteLength > limits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      data.byteLength,
      limits.maxGeneratedXmlBytes,
      part,
    );
  }
  return { data, patchBytes, patchCount: patches.length };
}

export function patchXlsxWorksheetStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetStructurePatchResult {
  let data: Uint8Array = bytes.slice();
  let patchBytes = 0;
  let patchCount = 0;
  for (const request of requested) {
    const result = patchOne(data, request, limits, part);
    data = result.data;
    patchBytes += result.patchBytes;
    patchCount += result.patchCount;
    if (patchBytes > limits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        limits.maxPatchBytes,
        part,
      );
    }
    if (patchCount > limits.maxPatchCount) {
      writeLimitFailure(
        'maxPatchCount',
        patchCount,
        limits.maxPatchCount,
        part,
      );
    }
  }
  return { data, patchBytes, patchCount };
}
