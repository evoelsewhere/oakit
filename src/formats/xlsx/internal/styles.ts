import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxStyle } from '../types';
import { xlsxBuiltinNumberFormatCode } from './number-format';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import { parseXlsxStyleBorder } from './style-border';
import { parseXlsxStyleFont } from './style-font';
import { parseXlsxStyleFill } from './style-fill';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

export interface XlsxCellXf {
  normalizedStyle: number;
  numberFormat?: string;
}

export interface XlsxStyleTable {
  cellXfs: readonly XlsxCellXf[];
  part: string | null;
  styles: readonly XlsxStyle[];
}

export const EMPTY_XLSX_STYLE_TABLE: XlsxStyleTable = Object.freeze({
  cellXfs: Object.freeze([]),
  part: null,
  styles: Object.freeze([]),
});

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const items: unknown[] = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of items) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function rootEntry(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
): { node: XmlRecord; prefix: string } {
  const document = value as unknown as XmlRecord;
  const entries = Object.entries(document);
  if (entries.length !== 1) {
    structureFailure('Styles root is missing', part);
  }
  const [qualifiedName, sourceNode] = entries[0]!;
  const node = record(sourceNode);
  const [first, second] = qualifiedName.split(':') as [string, string?];
  const prefix = second === undefined ? '' : first;
  if (!node || (second ?? first) !== 'styleSheet') {
    structureFailure('Styles root is missing', part);
  }
  const namespace = attributes(node)[prefix ? `xmlns:${prefix}` : 'xmlns'];
  if (namespace !== XLSX_SPREADSHEET_NAMESPACES[dialect]) {
    structureFailure('Styles root has the wrong namespace', part);
  }
  return { node, prefix };
}

function child(node: XmlRecord, prefix: string, localName: string): unknown {
  return node[prefix ? `${prefix}:${localName}` : localName];
}

function unsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) {
    valueFailure(message, part);
  }
  return parsed;
}

function collection(
  root: XmlRecord,
  prefix: string,
  collectionName: string,
  itemName: string,
  part: string,
  required: boolean,
): XmlRecord[] {
  const source = record(child(root, prefix, collectionName));
  if (!source) {
    if (!required) return [];
    structureFailure(`Styles ${collectionName} collection is missing`, part);
  }
  const items = records(child(source, prefix, itemName));
  if (!items) {
    structureFailure(`Styles ${collectionName} collection is invalid`, part);
  }
  const count = unsignedInteger(
    attributes(source).count,
    `Styles ${collectionName} count is invalid`,
    part,
  );
  if (count !== items.length) {
    structureFailure(`Styles ${collectionName} count does not match`, part);
  }
  if (required && items.length === 0) {
    structureFailure(`Styles ${collectionName} collection is empty`, part);
  }
  return items;
}

function customNumberFormats(
  root: XmlRecord,
  prefix: string,
  part: string,
): Map<number, string> {
  const values = new Map<number, string>();
  for (const item of collection(
    root,
    prefix,
    'numFmts',
    'numFmt',
    part,
    false,
  )) {
    const attrs = attributes(item);
    const id = unsignedInteger(
      attrs.numFmtId,
      'Styles number-format ID is invalid',
      part,
    );
    if (id < 164) {
      valueFailure('Styles custom number-format ID is reserved', part);
    }
    if (typeof attrs.formatCode !== 'string' || attrs.formatCode.length === 0) {
      valueFailure('Styles number-format code is invalid', part);
    }
    if (values.has(id)) {
      structureFailure('Styles contain a duplicate number-format ID', part);
    }
    values.set(id, attrs.formatCode);
  }
  return values;
}

function collectionCount(
  root: XmlRecord,
  prefix: string,
  name: 'cellStyleXfs',
  item: 'border' | 'fill' | 'font' | 'xf',
  part: string,
): number {
  return collection(root, prefix, name, item, part, true).length;
}

function referencedIndex(
  value: unknown,
  count: number,
  message: string,
  part: string,
): number {
  const index = value === undefined ? 0 : unsignedInteger(value, message, part);
  if (index >= count) valueFailure(message, part);
  return index;
}

function numberFormat(
  id: number,
  custom: ReadonlyMap<number, string>,
  part: string,
): string | undefined {
  if (id >= 164) {
    const code = custom.get(id);
    if (code === undefined) {
      valueFailure('Styles XF references a missing custom number format', part);
    }
    return code;
  }
  const code = xlsxBuiltinNumberFormatCode(id);
  if (code === undefined) {
    throw new XlsxParseError({
      code: 'unsupported-feature',
      message: 'Styles XF uses a locale-dependent built-in number format',
      part,
      severity: 'error',
    });
  }
  return code === 'General' ? undefined : code;
}

export function parseXlsxStylePart(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
  limits: ResolvedXlsxResourceLimits,
): XlsxStyleTable {
  const { node: root, prefix } = rootEntry(value, dialect, part);
  const custom = customNumberFormats(root, prefix, part);
  const fonts = collection(root, prefix, 'fonts', 'font', part, true).map(
    (font) => parseXlsxStyleFont(font, prefix, part),
  );
  const fills = collection(root, prefix, 'fills', 'fill', part, true).map(
    (fill) => parseXlsxStyleFill(fill, prefix, part),
  );
  const borders = collection(root, prefix, 'borders', 'border', part, true).map(
    (border) => parseXlsxStyleBorder(border, prefix, part),
  );
  const baseXfCount = collectionCount(root, prefix, 'cellStyleXfs', 'xf', part);
  const xfs = collection(root, prefix, 'cellXfs', 'xf', part, true);
  const totalStyles = custom.size + xfs.length;
  if (totalStyles > limits.maxStyles) {
    throw new XlsxResourceLimitError(
      'maxStyles',
      totalStyles,
      limits.maxStyles,
      part,
    );
  }

  const styles: XlsxStyle[] = [];
  const cellXfs: XlsxCellXf[] = [];
  const normalizedStyles = new Map<string, number>();
  for (const xf of xfs) {
    const attrs = attributes(xf);
    const numFmtId =
      attrs.numFmtId === undefined
        ? 0
        : unsignedInteger(
            attrs.numFmtId,
            'Styles XF number-format ID is invalid',
            part,
          );
    const fontId = referencedIndex(
      attrs.fontId,
      fonts.length,
      'Styles XF font reference is invalid',
      part,
    );
    const fillId = referencedIndex(
      attrs.fillId,
      fills.length,
      'Styles XF fill reference is invalid',
      part,
    );
    const borderId = referencedIndex(
      attrs.borderId,
      borders.length,
      'Styles XF border reference is invalid',
      part,
    );
    referencedIndex(
      attrs.xfId,
      baseXfCount,
      'Styles XF base-style reference is invalid',
      part,
    );
    const code = numberFormat(numFmtId, custom, part);
    const border = borders[borderId]!;
    const font = fonts[fontId]!;
    const fill = fills[fillId]!;
    const defaultFill =
      fill.kind === 'pattern' &&
      fill.pattern === 'none' &&
      fill.foregroundColor === undefined &&
      fill.backgroundColor === undefined;
    const style: XlsxStyle = {
      ...(Object.keys(border).length === 0 ? {} : { border }),
      ...(defaultFill ? {} : { fill }),
      ...(Object.keys(font).length === 0 ? {} : { font }),
      ...(code === undefined ? {} : { numberFormat: code }),
    };
    const styleKey = JSON.stringify(style);
    let normalizedStyle = normalizedStyles.get(styleKey);
    if (normalizedStyle === undefined) {
      normalizedStyle = styles.length;
      normalizedStyles.set(styleKey, normalizedStyle);
      styles.push(Object.freeze(style));
    }
    cellXfs.push(
      Object.freeze({
        normalizedStyle,
        ...(code === undefined ? {} : { numberFormat: code }),
      }),
    );
  }
  return Object.freeze({
    cellXfs: Object.freeze(cellXfs),
    part,
    styles: Object.freeze(styles),
  });
}

function stylesContentType(): string {
  return 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
}

function stylesRelationshipType(
  dialect: XlsxWorkbookDiscovery['dialect'],
): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/officeDocument/relationships/styles'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
}

export async function loadXlsxStyles(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxStyleTable> {
  const relationshipPart = getXlsxRelationshipPartName(discovery.part);
  const relationshipXml = await reader.readXml(relationshipPart, {
    required: true,
  });
  const relationships = parseXlsxRelationships(
    relationshipXml,
    discovery.part,
    limits.maxRelationships,
  );
  const relationshipType = stylesRelationshipType(discovery.dialect);
  const candidates = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  if (candidates.length === 0) return EMPTY_XLSX_STYLE_TABLE;
  if (candidates.length !== 1) {
    structureFailure(
      'Workbook contains multiple styles relationships',
      relationshipPart,
    );
  }
  const relationship = candidates[0]!;
  if (relationship.mode !== 'internal') {
    throw new XlsxParseError({
      code: 'invalid-relationship-target',
      message: 'Workbook styles relationship must be internal',
      part: relationshipPart,
      relationshipType: relationship.type,
      severity: 'error',
    });
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    stylesContentType()
  ) {
    structureFailure(
      'Workbook styles target has the wrong content type',
      relationship.target,
    );
  }
  const value = await reader.readXml(relationship.target, { required: true });
  return parseXlsxStylePart(
    value,
    discovery.dialect,
    relationship.target,
    limits,
  );
}
