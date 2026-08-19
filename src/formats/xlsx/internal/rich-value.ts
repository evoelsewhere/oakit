import { decodeXmlEntities } from '../../../common/text/html';
import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxRichValue,
  XlsxRichValueField,
  XlsxRichValueScalar,
} from '../types';
import type { XlsxCellMetadataBudget } from './cell-metadata';
import { XlsxPartReader } from './part-reader';
import type { XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

interface RichValueKey {
  name: string;
  type: XlsxRichValueField['type'];
}

interface RichValueStructure {
  keys: RichValueKey[];
  type: string;
}

export interface XlsxRichValueRegistry {
  part: string | null;
  values: readonly XlsxRichValue[];
}

export const EMPTY_XLSX_RICH_VALUES: XlsxRichValueRegistry = Object.freeze({
  part: null,
  values: Object.freeze([]),
});

function richNamespace(): string {
  return 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'missing-required-part',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const value of values) {
    const node = record(value);
    if (!node) return undefined;
    output.push(node);
  }
  return output;
}

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function sourcePrefix(name: string): string {
  const parts = name.split(':', 2);
  if (parts.length === 1) return '';
  const prefix = parts[0]!;
  return prefix.startsWith('ns_') ? prefix.slice(3) : prefix;
}

function namespaceFor(
  qualifiedName: string,
  node: unknown,
  inherited: Readonly<Record<string, string>>,
): string | undefined {
  const prefix = sourcePrefix(qualifiedName);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  const own = record(node);
  return (
    (own === undefined ? undefined : attributes(own)[declaration]) ??
    inherited[declaration]
  );
}

interface RootResult {
  attrs: Record<string, string>;
  node: XmlRecord;
}

function root(
  value: XmlLookupValue,
  expectedName: 'rvData' | 'rvStructures',
  part: string,
): RootResult {
  const entries = Object.entries(value).filter(
    ([name]) => localName(name) === expectedName,
  );
  if (entries.length !== 1) {
    fail(
      'invalid-document-structure',
      `Rich-value ${expectedName} root is missing or duplicated`,
      part,
    );
  }
  const [qualifiedName, rawNode] = entries[0]!;
  const node = record(rawNode);
  if (
    !node ||
    namespaceFor(qualifiedName, node, attributes(node)) !== richNamespace()
  ) {
    fail(
      'invalid-document-structure',
      `Rich-value ${expectedName} root has the wrong namespace`,
      part,
    );
  }
  return { attrs: attributes(node), node };
}

function children(
  node: XmlRecord,
  name: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): XmlRecord[] {
  const output: XmlRecord[] = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    if (localName(qualifiedName) !== name) continue;
    const values = records(value);
    if (!values) {
      fail(
        'invalid-document-structure',
        'Rich-value collection is invalid',
        part,
      );
    }
    for (const child of values) {
      if (namespaceFor(qualifiedName, child, inherited) !== richNamespace()) {
        fail(
          'invalid-document-structure',
          'Rich-value element has the wrong namespace',
          part,
        );
      }
      output.push(child);
    }
  }
  return output;
}

function elementValues(
  node: XmlRecord,
  name: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): unknown[] {
  const output: unknown[] = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    if (localName(qualifiedName) !== name) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const child of values) {
      if (namespaceFor(qualifiedName, child, inherited) !== richNamespace()) {
        fail(
          'invalid-document-structure',
          'Rich-value element has the wrong namespace',
          part,
        );
      }
      output.push(child);
    }
  }
  return output;
}

function scalar(node: unknown, message: string, part: string): string {
  if (typeof node === 'string') return decodeXmlEntities(node);
  const nodeRecord = record(node);
  if (!nodeRecord) fail('invalid-document-structure', message, part);
  const nested = Object.keys(nodeRecord).filter(
    (name) => name !== 'attrs' && name !== 'value',
  );
  if (nested.length !== 0) fail('invalid-document-structure', message, part);
  if (nodeRecord.value === undefined) return '';
  if (typeof nodeRecord.value !== 'string') {
    fail('invalid-document-value', message, part);
  }
  return decodeXmlEntities(nodeRecord.value);
}

function unsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) fail('invalid-document-value', message, part);
  return parsed;
}

function finiteNumber(value: string, message: string, part: string): number {
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) fail('invalid-document-value', message, part);
  return Object.is(parsed, -0) ? 0 : parsed;
}

function boolean(value: string, message: string, part: string): boolean {
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function assertCount(
  value: string | undefined,
  actual: number,
  message: string,
  part: string,
): void {
  if (value === undefined || unsignedInteger(value, message, part) !== actual) {
    fail('invalid-document-structure', message, part);
  }
}

function consumeRecord(
  budget: XlsxCellMetadataBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  amount = 1,
): void {
  const actual = budget.records + amount;
  if (!Number.isSafeInteger(actual) || actual > limits.maxMetadataRecords) {
    throw new XlsxResourceLimitError(
      'maxMetadataRecords',
      actual,
      limits.maxMetadataRecords,
      part,
    );
  }
  budget.records = actual;
}

function keyType(
  value: string | undefined,
  part: string,
): XlsxRichValueField['type'] {
  const type = value ?? 's';
  if (!['a', 'b', 'd', 'e', 'i', 'r', 's', 'spb'].includes(type)) {
    fail('invalid-document-value', 'Rich-value key type is invalid', part);
  }
  return type as XlsxRichValueField['type'];
}

function parseStructures(
  value: XmlLookupValue,
  part: string,
  budget: XlsxCellMetadataBudget,
  limits: ResolvedXlsxResourceLimits,
): RichValueStructure[] {
  const definition = root(value, 'rvStructures', part);
  const structures = children(definition.node, 's', definition.attrs, part);
  assertCount(
    attributes(definition.node).count,
    structures.length,
    'Rich-value structure count does not match',
    part,
  );
  return structures.map((structure) => {
    const attrs = attributes(structure);
    if (!attrs.t) {
      fail(
        'invalid-document-value',
        'Rich-value structure type is invalid',
        part,
      );
    }
    const keys = children(structure, 'k', definition.attrs, part).map((key) => {
      consumeRecord(budget, limits, part);
      const keyAttrs = attributes(key);
      if (!keyAttrs.n) {
        fail('invalid-document-value', 'Rich-value key name is invalid', part);
      }
      return {
        name: decodeXmlEntities(keyAttrs.n),
        type: keyType(keyAttrs.t, part),
      };
    });
    if (keys.length === 0) {
      fail(
        'invalid-document-structure',
        'Rich-value structure has no keys',
        part,
      );
    }
    return { keys, type: decodeXmlEntities(attrs.t) };
  });
}

function sensitiveKey(name: string): boolean {
  return /(?:^%|address|crid|identifier|license|provider|source|url|(?:^|_|entity|record|service)id$)/iu.test(
    name,
  );
}

function scalarValue(
  raw: string,
  type: XlsxRichValueField['type'],
  omitted: boolean,
  part: string,
): XlsxRichValueScalar {
  let parsed: XlsxRichValueScalar;
  if (type === 'b') {
    parsed = {
      kind: 'boolean',
      value: boolean(raw, 'Rich-value boolean is invalid', part),
    };
  } else if (type === 'd') {
    parsed = {
      kind: 'number',
      value: finiteNumber(raw, 'Rich-value number is invalid', part),
    };
  } else if (type === 'e') {
    if (!/^#[A-Z0-9/?!]+$/u.test(raw)) {
      fail('invalid-document-value', 'Rich-value error is invalid', part);
    }
    parsed = { code: raw, kind: 'error' };
  } else if (type === 'i') {
    parsed = {
      kind: 'integer',
      value: unsignedInteger(raw, 'Rich-value integer is invalid', part),
    };
  } else if (type === 'r') {
    parsed = {
      kind: 'rich-value-index',
      value: unsignedInteger(raw, 'Rich-value index is invalid', part),
    };
  } else if (type === 'a') {
    parsed = {
      kind: 'array-index',
      value: unsignedInteger(raw, 'Rich-value array index is invalid', part),
    };
  } else if (type === 'spb') {
    unsignedInteger(raw, 'Rich-value property-bag index is invalid', part);
    parsed = { kind: 'omitted' };
  } else {
    parsed = { kind: 'text', value: raw };
  }
  return omitted ? { kind: 'omitted' } : parsed;
}

function fallbackValue(node: unknown, part: string): XlsxRichValueScalar {
  const raw = scalar(node, 'Rich-value fallback is invalid', part);
  const nodeRecord = record(node);
  const type =
    nodeRecord === undefined ? 's' : (attributes(nodeRecord).t ?? 's');
  if (type === 's') return { kind: 'text', value: raw };
  if (type === 'b') {
    return {
      kind: 'boolean',
      value: boolean(raw, 'Rich-value fallback is invalid', part),
    };
  }
  if (type === 'n') {
    return {
      kind: 'number',
      value: finiteNumber(raw, 'Rich-value fallback is invalid', part),
    };
  }
  if (type === 'e' && /^#[A-Z0-9/?!]+$/u.test(raw)) {
    return { code: raw, kind: 'error' };
  }
  fail('invalid-document-value', 'Rich-value fallback type is invalid', part);
}

function parseData(
  value: XmlLookupValue,
  part: string,
  structures: readonly RichValueStructure[],
  budget: XlsxCellMetadataBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxRichValue[] {
  const definition = root(value, 'rvData', part);
  const values = children(definition.node, 'rv', definition.attrs, part);
  assertCount(
    attributes(definition.node).count,
    values.length,
    'Rich-value data count does not match',
    part,
  );
  const output = values.map((richValue) => {
    const structureIndex = unsignedInteger(
      attributes(richValue).s,
      'Rich-value structure reference is invalid',
      part,
    );
    const structure = structures[structureIndex];
    if (!structure) {
      fail(
        'invalid-document-value',
        'Rich-value structure reference is invalid',
        part,
      );
    }
    const rawValues = elementValues(richValue, 'v', definition.attrs, part);
    if (rawValues.length !== structure.keys.length) {
      fail(
        'invalid-document-structure',
        'Rich-value field count does not match',
        part,
      );
    }
    const fields = rawValues.map((node, index) => {
      consumeRecord(budget, limits, part);
      const key = structure.keys[index]!;
      const omitted = sensitiveKey(key.name);
      return {
        name: key.name,
        type: key.type,
        value: scalarValue(
          scalar(node, 'Rich-value field is invalid', part),
          key.type,
          omitted,
          part,
        ),
      };
    });
    const fallbacks = elementValues(richValue, 'fb', definition.attrs, part);
    if (fallbacks.length > 1) {
      fail(
        'invalid-document-structure',
        'Rich-value fallback is duplicated',
        part,
      );
    }
    return {
      ...(fallbacks[0] === undefined
        ? {}
        : { fallback: fallbackValue(fallbacks[0], part) }),
      fields,
      sourceDataOmitted: fields.some((field) => field.value.kind === 'omitted'),
      type: structure.type,
    };
  });
  for (const richValue of output) {
    for (const field of richValue.fields) {
      if (
        field.value.kind === 'rich-value-index' &&
        field.value.value >= output.length
      ) {
        fail('invalid-document-value', 'Rich-value reference is invalid', part);
      }
    }
  }
  return output;
}

function target(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  relationshipType: string,
  contentType: string,
  discovery: XlsxWorkbookDiscovery,
): string | undefined {
  const matches = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  if (matches.length > 1) {
    fail(
      'invalid-document-structure',
      'Rich-value relationship is duplicated',
      discovery.part,
    );
  }
  const relationship = matches[0];
  if (!relationship) return undefined;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Rich-value relationship must be internal',
      discovery.part,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !== contentType
  ) {
    fail(
      'invalid-document-structure',
      'Rich-value target has the wrong content type',
      relationship.target,
    );
  }
  return relationship.target;
}

export async function loadXlsxRichValues(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxCellMetadataBudget,
): Promise<XlsxRichValueRegistry> {
  const structurePart = target(
    relationships,
    'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure',
    'application/vnd.ms-excel.rdrichvaluestructure+xml',
    discovery,
  );
  const dataPart = target(
    relationships,
    'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue',
    'application/vnd.ms-excel.rdrichvalue+xml',
    discovery,
  );
  if (!structurePart && !dataPart) return EMPTY_XLSX_RICH_VALUES;
  if (!structurePart || !dataPart) {
    fail(
      'missing-required-part',
      'Rich-value data and structure parts must both exist',
      discovery.part,
    );
  }
  return parseXlsxRichValueParts(
    await reader.readXml(structurePart, { required: true }),
    await reader.readXml(dataPart, { required: true }),
    structurePart,
    dataPart,
    limits,
    budget,
  );
}

export function parseXlsxRichValueParts(
  structureValue: XmlLookupValue,
  dataValue: XmlLookupValue,
  structurePart: string,
  dataPart: string,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxCellMetadataBudget,
): XlsxRichValueRegistry {
  const structures = parseStructures(
    structureValue,
    structurePart,
    budget,
    limits,
  );
  return {
    part: dataPart,
    values: parseData(dataValue, dataPart, structures, budget, limits),
  };
}

function consumeText(
  budget: { textCharacters: number },
  value: string,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget.textCharacters + value.length;
  if (!Number.isSafeInteger(actual) || actual > limits.maxTextCharacters) {
    throw new XlsxResourceLimitError(
      'maxTextCharacters',
      actual,
      limits.maxTextCharacters,
      part,
    );
  }
  budget.textCharacters = actual;
}

export function cloneXlsxRichValueForOutput(
  value: XlsxRichValue,
  metadataBudget: XlsxCellMetadataBudget,
  textBudget: { textCharacters: number },
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxRichValue {
  consumeRecord(metadataBudget, limits, part, value.fields.length);
  consumeText(textBudget, value.type, limits, part);
  const cloneScalar = (scalar: XlsxRichValueScalar): XlsxRichValueScalar => {
    if (scalar.kind === 'text')
      consumeText(textBudget, scalar.value, limits, part);
    if (scalar.kind === 'error')
      consumeText(textBudget, scalar.code, limits, part);
    return { ...scalar };
  };
  return {
    ...(value.fallback === undefined
      ? {}
      : { fallback: cloneScalar(value.fallback) }),
    fields: value.fields.map((field) => {
      consumeText(textBudget, field.name, limits, part);
      return { ...field, value: cloneScalar(field.value) };
    }),
    sourceDataOmitted: value.sourceDataOmitted,
    type: value.type,
  };
}
