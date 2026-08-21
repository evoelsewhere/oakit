import { decodeXmlEntities } from '../../../common/text/html';
import { escapeXmlText } from '../writer/xml';
import { unsupportedPptxEdit } from './patch-error';
import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditableGraphicFrameXml,
  resolvePptxEditableShapeXml,
} from './shape-range';
import type { PptxRoundTripReplaceTextOperation } from './types';

const OFFICE_ESCAPE_PATTERN = /_x[0-9a-f]{4}_/i;

interface PptxTextPatchOwner {
  drawingPrefix: string;
  markupPrefix: string | undefined;
  presentationPrefix: string;
  xml: string;
}

interface PptxXmlRange {
  start: number;
  xml: string;
}

function assertCompatibleTextOwner(owner: PptxTextPatchOwner): void {
  if (
    (owner.markupPrefix !== undefined &&
      pptxShapeHasElement(
        owner.xml,
        qualifiedPptxName(owner.markupPrefix, 'AlternateContent'),
      )) ||
    pptxShapeHasElement(
      owner.xml,
      qualifiedPptxName(owner.presentationPrefix, 'extLst'),
    ) ||
    pptxShapeHasElement(
      owner.xml,
      qualifiedPptxName(owner.drawingPrefix, 'extLst'),
    )
  ) {
    unsupportedPptxEdit(
      'PowerPoint text edit target contains unsupported compatibility markup',
    );
  }
}

function patchPlainTextTarget(
  xml: string,
  target: string,
  targetStart: number,
  drawingPrefix: string,
  operation: PptxRoundTripReplaceTextOperation,
): string {
  if (
    pptxShapeHasElement(target, qualifiedPptxName(drawingPrefix, 'br')) ||
    pptxShapeHasElement(target, qualifiedPptxName(drawingPrefix, 'fld'))
  ) {
    unsupportedPptxEdit(
      'PowerPoint text edit target must contain one plain text run',
    );
  }

  const textName = qualifiedPptxName(drawingPrefix, 't');
  const escapedTextName = escapePptxXmlPattern(textName);
  const textPattern = new RegExp(
    `<${escapedTextName}((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*>([^<]*)<\\/${escapedTextName}\\s*>`,
    'g',
  );
  const matches = [...target.matchAll(textPattern)];
  if (matches.length !== 1) {
    unsupportedPptxEdit(
      'PowerPoint text edit target must contain exactly one text node',
    );
  }
  const match = matches[0] as RegExpMatchArray;
  const sourceText = decodeXmlEntities(match[2] as string);
  if (
    OFFICE_ESCAPE_PATTERN.test(sourceText) ||
    sourceText !== operation.expectedText
  ) {
    unsupportedPptxEdit(
      'PowerPoint text edit source XML does not match its preview precondition',
    );
  }
  const originalAttributes = match[1] as string;
  const attributesWithoutSpace = originalAttributes.replace(
    /\s+xml:space\s*=\s*(?:"[^"]*"|'[^']*')/gi,
    '',
  );
  const replacement = `<${textName}${attributesWithoutSpace} xml:space="preserve">${escapeXmlText(operation.value)}</${textName}>`;
  const matchStart = targetStart + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacement}${xml.slice(matchEnd)}`;
}

function xmlElementRanges(xml: string, qualifiedName: string): PptxXmlRange[] {
  const escapedName = escapePptxXmlPattern(qualifiedName);
  const pattern = new RegExp(
    `<${escapedName}(?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*\\s*>[\\s\\S]*?<\\/${escapedName}\\s*>`,
    'g',
  );
  return [...xml.matchAll(pattern)].map((match) => {
    const start = match.index ?? 0;
    return { start, xml: match[0] };
  });
}

function indexedRange(
  xml: string,
  qualifiedName: string,
  index: number,
  description: string,
): PptxXmlRange {
  const ranges = xmlElementRanges(xml, qualifiedName);
  const range = ranges[index];
  if (range === undefined) {
    unsupportedPptxEdit(
      `PowerPoint text edit target has no ${description} ${index + 1}`,
    );
  }
  return range;
}

export function patchPptxShapeTextXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripReplaceTextOperation,
): string {
  const { drawingPrefix, markupPrefix, presentationPrefix, range, shape } =
    resolvePptxEditableShapeXml(xml, shapeId);
  assertCompatibleTextOwner({
    drawingPrefix,
    markupPrefix,
    presentationPrefix,
    xml: shape,
  });
  return patchPlainTextTarget(
    xml,
    shape,
    range.start,
    drawingPrefix,
    operation,
  );
}

export function patchPptxTableCellTextXml(
  xml: string,
  shapeId: string,
  rowIndex: number,
  columnIndex: number,
  operation: PptxRoundTripReplaceTextOperation,
): string {
  if (
    !Number.isSafeInteger(rowIndex) ||
    rowIndex < 0 ||
    !Number.isSafeInteger(columnIndex) ||
    columnIndex < 0
  ) {
    unsupportedPptxEdit('PowerPoint table text edit target index is unsafe');
  }
  const frame = resolvePptxEditableGraphicFrameXml(xml, shapeId);
  assertCompatibleTextOwner({
    drawingPrefix: frame.drawingPrefix,
    markupPrefix: frame.markupPrefix,
    presentationPrefix: frame.presentationPrefix,
    xml: frame.shape,
  });
  const tableName = qualifiedPptxName(frame.drawingPrefix, 'tbl');
  const tables = xmlElementRanges(frame.shape, tableName);
  if (tables.length !== 1) {
    unsupportedPptxEdit(
      'PowerPoint table text edit requires exactly one native table',
    );
  }
  const table = tables[0] as PptxXmlRange;
  const row = indexedRange(
    table.xml,
    qualifiedPptxName(frame.drawingPrefix, 'tr'),
    rowIndex,
    'table row',
  );
  const cell = indexedRange(
    row.xml,
    qualifiedPptxName(frame.drawingPrefix, 'tc'),
    columnIndex,
    'table cell',
  );
  assertCompatibleTextOwner({
    drawingPrefix: frame.drawingPrefix,
    markupPrefix: frame.markupPrefix,
    presentationPrefix: frame.presentationPrefix,
    xml: cell.xml,
  });
  return patchPlainTextTarget(
    xml,
    cell.xml,
    frame.range.start + table.start + row.start + cell.start,
    frame.drawingPrefix,
    operation,
  );
}
