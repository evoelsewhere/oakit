import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
} from '../../src/formats/xlsx/internal/resource-limits';
import {
  EMPTY_XLSX_STYLE_TABLE,
  loadXlsxStyles,
  parseXlsxStylePart,
} from '../../src/formats/xlsx/internal/styles';
import {
  discoverXlsxWorkbook,
  XLSX_SPREADSHEET_NAMESPACES,
} from '../../src/formats/xlsx/internal/workbook-discovery';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
  type XlsxBlackBoxOverrides,
} from '../black-box/xlsx-package';

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits> = {},
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

async function load(
  overrides: XlsxBlackBoxOverrides = {},
  limitOverrides: Partial<ResolvedXlsxResourceLimits> = {},
) {
  const resolved = limits(limitOverrides);
  const zip = await JSZip.loadAsync(await createIndependentXlsx(overrides));
  const reader = new XlsxPartReader(zip, [], resolved);
  const discovery = await discoverXlsxWorkbook(reader, resolved);
  return loadXlsxStyles(discovery, reader, resolved);
}

async function capture(
  overrides: XlsxBlackBoxOverrides,
): Promise<XlsxParseError> {
  try {
    await load(overrides);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected styles parsing to fail');
}

function styleSheet(body: string, namespace = XLSX_SPREADSHEET_NS): string {
  return `<styleSheet xmlns="${namespace}">${body}</styleSheet>`;
}

const CORE = `
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>`;

const STRICT_OFFICE_REL_NS =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';

function parseTree(
  value: unknown,
  dialect: 'strict' | 'transitional' = 'transitional',
) {
  return parseXlsxStylePart(
    value as XmlLookupValue,
    dialect,
    'xl/styles.xml',
    limits(),
  );
}

function captureTree(tree: unknown): XlsxParseError {
  try {
    parseTree(tree);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected styles tree parsing to fail');
}

function contentTypes(stylesType: string): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="${stylesType}"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  </Types>`;
}

describe('XLSX styles table', () => {
  it('parses built-in and custom number formats in authored XF order', async () => {
    const result = await load({
      'xl/styles.xml': styleSheet(`${CORE}
        <numFmts count="2">
          <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
          <numFmt numFmtId="4294967295" formatCode="[h]:mm:ss"/>
        </numFmts>
        <cellXfs count="5">
          <xf/>
          <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
          <xf numFmtId="164"/>
          <xf numFmtId="4294967295"/>
          <xf numFmtId="14"/>
        </cellXfs>`),
    });

    expect(result).toEqual({
      cellXfs: [
        { normalizedStyle: 0 },
        { normalizedStyle: 1, numberFormat: 'mm-dd-yy' },
        { normalizedStyle: 2, numberFormat: 'yyyy-mm-dd' },
        { normalizedStyle: 3, numberFormat: '[h]:mm:ss' },
        { normalizedStyle: 1, numberFormat: 'mm-dd-yy' },
      ],
      part: 'xl/styles.xml',
      styles: [
        {},
        { numberFormat: 'mm-dd-yy' },
        { numberFormat: 'yyyy-mm-dd' },
        { numberFormat: '[h]:mm:ss' },
      ],
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.cellXfs)).toBe(true);
    expect(Object.isFrozen(result.cellXfs[0])).toBe(true);
    expect(Object.isFrozen(result.styles)).toBe(true);
    expect(Object.isFrozen(result.styles[0])).toBe(true);
  });

  it('parses prefixed Strict styles through Strict package relationships', async () => {
    const namespace = XLSX_SPREADSHEET_NAMESPACES.strict;
    const relationshipBase = `${STRICT_OFFICE_REL_NS}/`;
    const result = await load({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="main" Type="${relationshipBase}officeDocument" Target="xl/workbook.xml"/>
      </Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="styles" Type="${relationshipBase}styles" Target="styles.xml"/>
      </Relationships>`,
      'xl/styles.xml': `<s:styleSheet xmlns:s="${namespace}">
        <s:fonts count="1"><s:font/></s:fonts>
        <s:fills count="1"><s:fill/></s:fills>
        <s:borders count="1"><s:border/></s:borders>
        <s:cellStyleXfs count="1"><s:xf/></s:cellStyleXfs>
        <s:cellXfs count="1"><s:xf numFmtId="14"/></s:cellXfs>
      </s:styleSheet>`,
      'xl/workbook.xml': `<s:workbook xmlns:s="${namespace}" xmlns:r="${STRICT_OFFICE_REL_NS}">
        <s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets>
      </s:workbook>`,
    });

    expect(result).toMatchObject({
      cellXfs: [{ normalizedStyle: 0, numberFormat: 'mm-dd-yy' }],
      styles: [{ numberFormat: 'mm-dd-yy' }],
    });
  });

  it('returns the immutable empty table when the workbook has no styles relationship', async () => {
    const result = await load({
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
      </Relationships>`,
      'xl/styles.xml': null,
    });

    expect(result).toBe(EMPTY_XLSX_STYLE_TABLE);
    expect(Object.isFrozen(result)).toBe(true);
  });

  it('resolves font references into semantic styles and deduplicates them', async () => {
    const result = await load({
      'xl/styles.xml': styleSheet(`
        <fonts count="2">
          <font/>
          <font><b/><name val="Aptos"/><color theme="4" tint=".25"/></font>
        </fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf/></cellStyleXfs>
        <cellXfs count="4">
          <xf/>
          <xf fontId="1"/>
          <xf fontId="1"/>
          <xf fontId="1" numFmtId="14"/>
        </cellXfs>`),
    });

    expect(result).toEqual({
      cellXfs: [
        { normalizedStyle: 0 },
        { normalizedStyle: 1 },
        { normalizedStyle: 1 },
        { normalizedStyle: 2, numberFormat: 'mm-dd-yy' },
      ],
      part: 'xl/styles.xml',
      styles: [
        {},
        {
          font: {
            bold: true,
            color: { index: 4, kind: 'theme', tint: 0.25 },
            name: 'Aptos',
          },
        },
        {
          font: {
            bold: true,
            color: { index: 4, kind: 'theme', tint: 0.25 },
            name: 'Aptos',
          },
          numberFormat: 'mm-dd-yy',
        },
      ],
    });
    expect(Object.isFrozen(result.styles[1]?.font)).toBe(true);
  });

  it('accepts maxStyles exactly and rejects one over using formats plus XFs', async () => {
    const xml = styleSheet(`${CORE}
      <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
      <cellXfs count="2"><xf/><xf numFmtId="164"/></cellXfs>`);
    await expect(
      load({ 'xl/styles.xml': xml }, { maxStyles: 3 }),
    ).resolves.toMatchObject({
      styles: [{}, { numberFormat: '0.000' }],
    });
    await expect(
      load({ 'xl/styles.xml': xml }, { maxStyles: 2 }),
    ).rejects.toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxStyles',
      name: 'XlsxResourceLimitError',
      part: 'xl/styles.xml',
    });
  });

  it.each([
    [
      styleSheet(`${CORE}<cellXfs count="1"><xf numFmtId="5"/></cellXfs>`),
      'unsupported-feature',
      'Styles XF uses a locale-dependent built-in number format',
    ],
    [
      styleSheet(`${CORE}<cellXfs count="1"><xf numFmtId="164"/></cellXfs>`),
      'invalid-document-value',
      'Styles XF references a missing custom number format',
    ],
    [
      styleSheet(
        `${CORE}<numFmts count="1"><numFmt numFmtId="1" formatCode="x"/></numFmts><cellXfs count="1"><xf/></cellXfs>`,
      ),
      'invalid-document-value',
      'Styles custom number-format ID is reserved',
    ],
    [
      styleSheet(
        `${CORE}<numFmts count="1"><numFmt numFmtId="164" formatCode=""/></numFmts><cellXfs count="1"><xf/></cellXfs>`,
      ),
      'invalid-document-value',
      'Styles number-format code is invalid',
    ],
    [
      styleSheet(
        `${CORE}<numFmts count="2"><numFmt numFmtId="164" formatCode="x"/><numFmt numFmtId="164" formatCode="y"/></numFmts><cellXfs count="1"><xf/></cellXfs>`,
      ),
      'invalid-document-structure',
      'Styles contain a duplicate number-format ID',
    ],
    [
      styleSheet(
        `${CORE}<numFmts count="1"><numFmt numFmtId="164junk" formatCode="x"/></numFmts><cellXfs count="1"><xf/></cellXfs>`,
      ),
      'invalid-document-value',
      'Styles number-format ID is invalid',
    ],
    [
      styleSheet(
        `${CORE}<numFmts count="1"><numFmt numFmtId="4294967296" formatCode="x"/></numFmts><cellXfs count="1"><xf/></cellXfs>`,
      ),
      'invalid-document-value',
      'Styles number-format ID is invalid',
    ],
  ] as const)(
    'rejects invalid number format %#',
    async (xml, code, message) => {
      const error = await capture({ 'xl/styles.xml': xml });
      expect(error.diagnostic).toMatchObject({
        code,
        message,
        part: 'xl/styles.xml',
      });
    },
  );

  it.each([
    ['fontId', 1, 'Styles XF font reference is invalid'],
    ['fillId', 1, 'Styles XF fill reference is invalid'],
    ['borderId', 1, 'Styles XF border reference is invalid'],
    ['xfId', 1, 'Styles XF base-style reference is invalid'],
    ['numFmtId', -1, 'Styles XF number-format ID is invalid'],
  ] as const)(
    'rejects invalid XF %s reference',
    async (attribute, value, message) => {
      const error = await capture({
        'xl/styles.xml': styleSheet(
          `${CORE}<cellXfs count="1"><xf ${attribute}="${value}"/></cellXfs>`,
        ),
      });
      expect(error.diagnostic).toMatchObject({
        code: 'invalid-document-value',
        message,
      });
    },
  );

  it.each(['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs'] as const)(
    'requires a non-empty %s collection',
    async (name) => {
      const complete = `${CORE}<cellXfs count="1"><xf/></cellXfs>`;
      const without = complete.replace(
        new RegExp(`<${name}[^>]*>.*?</${name}>`),
        '',
      );
      const error = await capture({
        'xl/styles.xml': styleSheet(without),
      });
      expect(error.diagnostic.message).toBe(
        `Styles ${name} collection is missing`,
      );

      const empty = complete.replace(
        new RegExp(`<${name}[^>]*>.*?</${name}>`),
        `<${name} count="0"></${name}>`,
      );
      const emptyError = await capture({
        'xl/styles.xml': styleSheet(empty),
      });
      expect(emptyError.diagnostic.message).toBe(
        `Styles ${name} collection is empty`,
      );
    },
  );

  it('rejects mismatched and invalid collection counts', async () => {
    const mismatch = await capture({
      'xl/styles.xml': styleSheet(`${CORE}<cellXfs count="2"><xf/></cellXfs>`),
    });
    expect(mismatch.diagnostic.message).toBe(
      'Styles cellXfs count does not match',
    );

    const invalid = await capture({
      'xl/styles.xml': styleSheet(`${CORE}<cellXfs count="-1"><xf/></cellXfs>`),
    });
    expect(invalid.diagnostic.message).toBe('Styles cellXfs count is invalid');
  });

  it.each([
    ['<worksheet/>', 'Styles root is missing'],
    ['<styleSheet xmlns="urn:wrong"/>', 'Styles root has the wrong namespace'],
  ])('rejects invalid styles root %#', async (xml, message) => {
    const error = await capture({ 'xl/styles.xml': xml });
    expect(error.diagnostic.message).toBe(message);
  });

  it.each([
    [{}, 'Styles root is missing'],
    [{ styleSheet: 'not-an-element' }, 'Styles root is missing'],
    [{ styleSheet: {}, worksheet: {} }, 'Styles root is missing'],
  ])('rejects malformed in-memory styles root %#', (tree, message) => {
    expect(captureTree(tree).diagnostic).toMatchObject({ message });
  });

  it('rejects a non-element collection item', () => {
    const root = {
      attrs: { xmlns: XLSX_SPREADSHEET_NS },
      borders: { attrs: { count: '1' }, border: {} },
      cellStyleXfs: { attrs: { count: '1' }, xf: {} },
      cellXfs: { attrs: { count: '1' }, xf: {} },
      fills: { attrs: { count: '1' }, fill: {} },
      fonts: { attrs: { count: '1' }, font: 'not-an-element' },
    };

    expect(captureTree({ styleSheet: root }).diagnostic).toMatchObject({
      message: 'Styles fonts collection is invalid',
      part: 'xl/styles.xml',
    });
  });

  it('rejects duplicate, external, wrong-content-type, and missing styles targets', async () => {
    const relationships = (entries: string) =>
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">${entries}</Relationships>`;
    const styleRelationship = (id: string, extra = '') =>
      `<Relationship Id="${id}" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"${extra}/>`;

    await expect(
      load({
        'xl/_rels/workbook.xml.rels': relationships(
          `${styleRelationship('a')}${styleRelationship('b')}`,
        ),
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        message: 'Workbook contains multiple styles relationships',
      },
    });
    await expect(
      load({
        'xl/_rels/workbook.xml.rels': relationships(
          styleRelationship('a', ' TargetMode="External"'),
        ),
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-relationship-target',
        message: 'Workbook styles relationship must be internal',
      },
    });
    await expect(
      load({
        '[Content_Types].xml': contentTypes('application/xml'),
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        message: 'Workbook styles target has the wrong content type',
        part: 'xl/styles.xml',
      },
    });
    await expect(load({ 'xl/styles.xml': null })).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        part: 'xl/styles.xml',
      },
    });
    await expect(
      load({ 'xl/_rels/workbook.xml.rels': null }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        part: 'xl/_rels/workbook.xml.rels',
      },
    });
  });
});
