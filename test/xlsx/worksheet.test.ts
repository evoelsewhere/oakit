import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxSharedStringTable } from '../../src/formats/xlsx/internal/shared-strings';
import type { XlsxResolvedSheetSelection } from '../../src/formats/xlsx/internal/selection';
import {
  createXlsxWorksheetBudget,
  parseXlsxWorksheetPart,
  type XlsxWorksheetBudget,
} from '../../src/formats/xlsx/internal/worksheet';
import { XLSX_SPREADSHEET_NAMESPACES } from '../../src/formats/xlsx/internal/workbook-discovery';

const PART = 'xl/worksheets/sheet1.xml';
const TRANSITIONAL = XLSX_SPREADSHEET_NAMESPACES.transitional;
const STRICT = XLSX_SPREADSHEET_NAMESPACES.strict;
const EMPTY_STRINGS: XlsxSharedStringTable = { part: null, values: [] };
const SHARED_STRINGS: XlsxSharedStringTable = {
  part: 'xl/sharedStrings.xml',
  values: [
    { text: 'Shared' },
    {
      runs: [{ text: 'Rich' }, { text: ' shared' }],
      text: 'Rich shared',
    },
  ],
};

function worksheet(body: string, namespace: string = TRANSITIONAL): string {
  return `<worksheet xmlns="${namespace}">${body}</worksheet>`;
}

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits> = {},
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

async function parse(
  xml: string,
  options: {
    budget?: XlsxWorksheetBudget;
    dialect?: 'strict' | 'transitional';
    limits?: Partial<ResolvedXlsxResourceLimits>;
    part?: string;
    selection?: XlsxResolvedSheetSelection;
    strings?: XlsxSharedStringTable;
  } = {},
) {
  const part = options.part ?? PART;
  const strings = options.strings ?? EMPTY_STRINGS;
  const zip = new JSZip();
  zip.file(part, xml);
  const resolved = limits(options.limits);
  const reader = new XlsxPartReader(zip, [], resolved);
  return parseXlsxWorksheetPart(
    part,
    options.dialect ?? 'transitional',
    reader,
    resolved,
    strings,
    options.budget ?? createXlsxWorksheetBudget(strings),
    options.selection,
  );
}

async function captureParseError(
  xml: string,
  options: Parameters<typeof parse>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parse(xml, options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected worksheet parsing to fail');
}

describe('XLSX worksheet streaming', () => {
  it('parses sparse scalar, shared, inline, blank, and row metadata', async () => {
    const xml = worksheet(`
      <dimension ref="A1:XFD1048576"/>
      <sheetViews><sheetView workbookViewId="0"/></sheetViews>
      <sheetData>
        <row r="2" ht="20.5" hidden="true" outlineLevel="7" spans="1:20">
          <c r="A2" s="3"><v>42.5</v></c>
          <c><v>-0</v></c>
          <c r="C2" t="b"><v>1</v></c>
          <c r="D2" t="e"><v>#DIV/0!</v></c>
          <c r="E2" t="str"><v> formula text </v></c>
          <c r="F2" t="d"><v>2024-02-29T23:59:59Z</v></c>
          <c r="G2" t="s"><v>0</v></c>
          <c r="H2" t="s"><v>1</v></c>
          <c r="I2" t="inlineStr"><is><t xml:space="preserve"> inline </t></is></c>
          <c r="J2" t="inlineStr"><is>
            <r><rPr><b/><color rgb="FF000000"/></rPr><t>Rich</t></r>
            <r><t xml:space="default"> inline</t></r>
            <rPh sb="0" eb="4"><t>ignored</t></rPh>
            <phoneticPr fontId="1"/>
          </is></c>
          <c r="K2" s="4"/>
          <extLst><ext uri="urn:test"><foreign xmlns="urn:foreign"/></ext></extLst>
        </row>
        <row hidden="0" ht="0" outlineLevel="0">
          <c t="n"><v>7</v></c>
          <c t="inlineStr"><is/></c>
          <c t="str"/>
        </row>
      </sheetData>
      <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
      <extLst><ext uri="urn:test"><x:payload xmlns:x="urn:extension"/></ext></extLst>
    `);

    await expect(parse(xml, { strings: SHARED_STRINGS })).resolves.toEqual({
      rows: [
        {
          cells: [
            {
              address: 'A2',
              column: 1,
              content: {
                kind: 'value',
                value: { kind: 'number', value: 42.5 },
              },
              style: 3,
            },
            {
              address: 'B2',
              column: 2,
              content: { kind: 'value', value: { kind: 'number', value: 0 } },
            },
            {
              address: 'C2',
              column: 3,
              content: {
                kind: 'value',
                value: { kind: 'boolean', value: true },
              },
            },
            {
              address: 'D2',
              column: 4,
              content: {
                kind: 'value',
                value: { code: '#DIV/0!', kind: 'error' },
              },
            },
            {
              address: 'E2',
              column: 5,
              content: {
                kind: 'value',
                value: { kind: 'text', text: ' formula text ' },
              },
            },
            {
              address: 'F2',
              column: 6,
              content: {
                kind: 'value',
                value: {
                  kind: 'date',
                  normalized: '2024-02-29T23:59:59Z',
                  precision: 'date-time',
                  source: { kind: 'iso', value: '2024-02-29T23:59:59Z' },
                },
              },
            },
            {
              address: 'G2',
              column: 7,
              content: {
                kind: 'value',
                value: { kind: 'text', text: 'Shared' },
              },
            },
            {
              address: 'H2',
              column: 8,
              content: {
                kind: 'value',
                value: {
                  kind: 'text',
                  runs: [{ text: 'Rich' }, { text: ' shared' }],
                  text: 'Rich shared',
                },
              },
            },
            {
              address: 'I2',
              column: 9,
              content: {
                kind: 'value',
                value: { kind: 'text', text: ' inline ' },
              },
            },
            {
              address: 'J2',
              column: 10,
              content: {
                kind: 'value',
                value: {
                  kind: 'text',
                  runs: [{ text: 'Rich' }, { text: ' inline' }],
                  text: 'Rich inline',
                },
              },
            },
            {
              address: 'K2',
              column: 11,
              content: { kind: 'blank' },
              style: 4,
            },
          ],
          height: 20.5,
          hidden: true,
          index: 2,
          outlineLevel: 7,
        },
        {
          cells: [
            {
              address: 'A3',
              column: 1,
              content: { kind: 'value', value: { kind: 'number', value: 7 } },
            },
            {
              address: 'B3',
              column: 2,
              content: { kind: 'value', value: { kind: 'text', text: '' } },
            },
            { address: 'C3', column: 3, content: { kind: 'blank' } },
          ],
          height: 0,
          hidden: false,
          index: 3,
          outlineLevel: 0,
        },
      ],
    });
  });

  it('parses prefixed Strict worksheet elements', async () => {
    const xml = `<s:worksheet xmlns:s="${STRICT}">
      <s:sheetData><s:row><s:c t="inlineStr"><s:is><s:t>Strict</s:t></s:is></s:c></s:row></s:sheetData>
    </s:worksheet>`;

    await expect(parse(xml, { dialect: 'strict' })).resolves.toEqual({
      rows: [
        {
          cells: [
            {
              address: 'A1',
              column: 1,
              content: {
                kind: 'value',
                value: { kind: 'text', text: 'Strict' },
              },
            },
          ],
          index: 1,
        },
      ],
    });
  });

  it('accepts empty sheet data and whitespace-only structural text', async () => {
    await expect(
      parse(worksheet('\n<sheetData> \n </sheetData>\n')),
    ).resolves.toEqual({
      rows: [],
    });
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)('normalizes row hidden value %s', async (source, expected) => {
    const result = await parse(
      worksheet(`<sheetData><row hidden="${source}"/></sheetData>`),
    );
    expect(result.rows[0]?.hidden).toBe(expected);
  });

  it.each([
    ['.5', 0.5],
    ['.55', 0.55],
    ['1.', 1],
    ['12', 12],
    ['17976931348623157' + '0'.repeat(292), 1.7976931348623157e308],
  ] as const)('parses finite row height %s', async (source, expected) => {
    const result = await parse(
      worksheet(`<sheetData><row ht="${source}"/></sheetData>`),
    );
    expect(result.rows[0]?.height).toBe(expected);
  });

  it('preserves unsigned style boundaries and explicit unstyled blanks', async () => {
    const result = await parse(
      worksheet(
        '<sheetData><row><c s="0"/><c s="4294967295"/><c/></row></sheetData>',
      ),
    );
    expect(result.rows[0]?.cells).toEqual([
      { address: 'A1', column: 1, content: { kind: 'blank' }, style: 0 },
      {
        address: 'B1',
        column: 2,
        content: { kind: 'blank' },
        style: 0xffff_ffff,
      },
      { address: 'C1', column: 3, content: { kind: 'blank' } },
    ]);
    expect('style' in result.rows[0]!.cells[2]!).toBe(false);
  });

  it('ignores cell extension payload without treating nested foreign XML as cells', async () => {
    const result = await parse(
      worksheet(
        '<sheetData><row><c><extLst><ext uri="urn:test"><x:payload xmlns:x="urn:x"/></ext></extLst></c></row></sheetData>',
      ),
    );
    expect(result).toEqual({
      rows: [
        {
          cells: [{ address: 'A1', column: 1, content: { kind: 'blank' } }],
          index: 1,
        },
      ],
    });
  });

  it('initializes aggregate text and rich-run accounting from shared strings', () => {
    const table: XlsxSharedStringTable = {
      part: 'strings.xml',
      values: [
        { text: 'abc' },
        {
          phoneticRuns: [
            { end: 1, start: 0, text: 'xy' },
            { end: 2, start: 1, text: 'z' },
          ],
          runs: [{ text: 'd' }, { text: 'ef' }],
          text: 'def',
        },
      ],
    };

    expect(createXlsxWorksheetBudget(table)).toEqual({
      formulaCharacters: 0,
      formulaGroups: 0,
      returnedCells: 0,
      richTextRuns: 4,
      scannedCells: 0,
      textCharacters: 9,
    });
  });

  it('shares cell and text budgets across worksheets', async () => {
    const strings: XlsxSharedStringTable = {
      part: 'strings.xml',
      values: [{ text: 'A' }],
    };
    const budget = createXlsxWorksheetBudget(strings);
    const first = worksheet(
      '<sheetData><row><c t="s"><v>0</v></c></row></sheetData>',
    );
    const second = worksheet(
      '<sheetData><row><c t="str"><v>B</v></c></row></sheetData>',
    );

    await parse(first, {
      budget,
      limits: {
        maxReturnedCells: 2,
        maxScannedCells: 2,
        maxTextCharacters: 3,
      },
      strings,
    });
    await parse(second, {
      budget,
      limits: {
        maxReturnedCells: 2,
        maxScannedCells: 2,
        maxTextCharacters: 3,
      },
      part: 'xl/worksheets/sheet2.xml',
      strings,
    });

    expect(budget).toEqual({
      formulaCharacters: 0,
      formulaGroups: 0,
      returnedCells: 2,
      richTextRuns: 0,
      scannedCells: 2,
      textCharacters: 3,
    });
  });

  it('emits selected cells and authored row metadata without post-filtering', async () => {
    const budget = createXlsxWorksheetBudget(EMPTY_STRINGS);
    const selection: XlsxResolvedSheetSelection = {
      endRowPrefix: [2],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 2, row: 2 },
          reference: 'B1:B2',
          start: { column: 2, row: 1 },
        },
      ],
    };
    const result = await parse(
      worksheet(`<sheetData>
        <row r="1" hidden="true">
          <c r="A1" t="inlineStr"><is><t>outside</t></is></c>
          <c r="B1" t="inlineStr"><is><t>inside</t></is></c>
        </row>
        <row r="2" ht="12"><c r="A2"><v>7</v></c></row>
      </sheetData>`),
      { budget, selection },
    );

    expect(result).toEqual({
      rows: [
        {
          cells: [
            {
              address: 'B1',
              column: 2,
              content: {
                kind: 'value',
                value: { kind: 'text', text: 'inside' },
              },
            },
          ],
          hidden: true,
          index: 1,
        },
        { cells: [], height: 12, index: 2 },
      ],
    });
    expect(budget).toEqual({
      formulaCharacters: 0,
      formulaGroups: 0,
      returnedCells: 1,
      richTextRuns: 0,
      scannedCells: 3,
      textCharacters: 6,
    });
  });

  it('validates unreturned cells and charges scanned work independently', async () => {
    const selection: XlsxResolvedSheetSelection = {
      endRowPrefix: [1],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 1, row: 1 },
          reference: 'A1',
          start: { column: 1, row: 1 },
        },
      ],
    };
    const valid = worksheet(
      '<sheetData><row><c><v>1</v></c><c><v>2</v></c></row></sheetData>',
    );
    const budget = createXlsxWorksheetBudget(EMPTY_STRINGS);
    await expect(
      parse(valid, {
        budget,
        limits: { maxReturnedCells: 1, maxScannedCells: 2 },
        selection,
      }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    expect(budget.returnedCells).toBe(1);
    expect(budget.scannedCells).toBe(2);

    await expect(
      parse(valid, {
        limits: { maxReturnedCells: 1, maxScannedCells: 1 },
        selection,
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxScannedCells',
      name: 'XlsxResourceLimitError',
      part: PART,
    } satisfies Partial<XlsxResourceLimitError>);

    const invalid = worksheet(
      '<sheetData><row><c><v>1</v></c><c t="b"><v>2</v></c></row></sheetData>',
    );
    await expect(parse(invalid, { selection })).rejects.toMatchObject({
      diagnostic: {
        cell: 'B1',
        code: 'invalid-document-value',
        message: 'Cell boolean is invalid',
        part: PART,
      },
      name: 'XlsxParseError',
    });
  });

  it('bounds selected output separately from unreturned inline resources', async () => {
    const oneCell: XlsxResolvedSheetSelection = {
      endRowPrefix: [1],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 1, row: 1 },
          reference: 'A1',
          start: { column: 1, row: 1 },
        },
      ],
    };
    const bothCells: XlsxResolvedSheetSelection = {
      endRowPrefix: [1],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 2, row: 1 },
          reference: 'A1:B1',
          start: { column: 1, row: 1 },
        },
      ],
    };
    const cells = worksheet(
      '<sheetData><row><c/><c t="inlineStr"><is><r><t>A</t></r><r><t>B</t></r></is></c></row></sheetData>',
    );

    await expect(
      parse(cells, {
        limits: { maxReturnedCells: 1, maxRichTextRuns: 2 },
        selection: oneCell,
      }),
    ).resolves.toEqual({
      rows: [
        {
          cells: [{ address: 'A1', column: 1, content: { kind: 'blank' } }],
          index: 1,
        },
      ],
    });
    await expect(
      parse(cells, {
        limits: { maxReturnedCells: 1, maxRichTextRuns: 2 },
        selection: bothCells,
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxReturnedCells',
    });
    await expect(
      parse(cells, {
        limits: { maxReturnedCells: 1, maxRichTextRuns: 1 },
        selection: oneCell,
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxRichTextRuns',
    });

    const scalarText = worksheet(
      '<sheetData><row><c/><c t="str"><v>outside</v></c></row></sheetData>',
    );
    await expect(
      parse(scalarText, {
        limits: { maxReturnedCells: 1, maxTextCharacters: 1 },
        selection: oneCell,
      }),
    ).resolves.toEqual({
      rows: [
        {
          cells: [{ address: 'A1', column: 1, content: { kind: 'blank' } }],
          index: 1,
        },
      ],
    });
  });

  it.each([
    [
      '<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      'Worksheet root is missing',
      undefined,
    ],
    [
      worksheet('<sheetData/>', 'urn:wrong'),
      'Worksheet element has an unsupported namespace',
      undefined,
    ],
    [worksheet(''), 'Worksheet sheetData is missing', undefined],
    [
      worksheet('<sheetData/><sheetData/>'),
      'Worksheet contains duplicate sheetData elements',
      undefined,
    ],
    [
      worksheet('<sheetData>text</sheetData>'),
      'Worksheet text is outside a value or inline-string text element',
      undefined,
    ],
    [
      worksheet('<sheetData><c/></sheetData>'),
      'Worksheet element nesting is invalid',
      undefined,
    ],
    [
      worksheet('<sheetData><row><unknown/></row></sheetData>'),
      'Worksheet element nesting is invalid',
      undefined,
    ],
    [
      worksheet('<sheetData><row><c><unknown/></c></row></sheetData>'),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet('<sheetData><row><c><v><t>1</t></v></c></row></sheetData>'),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet(
        '<sheetData><row><c t="inlineStr"><is><row/></is></c></row></sheetData>',
      ),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet(
        '<sheetData><row><c t="inlineStr"><is><v>A</v></is></c></row></sheetData>',
      ),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet(
        '<sheetData><row><c t="inlineStr"><is><unknown/></is></c></row></sheetData>',
      ),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet(
        '<sheetData><row><c t="inlineStr"><is><r><unknown/></r></is></c></row></sheetData>',
      ),
      'Worksheet element nesting is invalid',
      'A1',
    ],
    [
      worksheet('<sheetData><row><c xmlns="urn:wrong"/></row></sheetData>'),
      'Worksheet element has an unsupported namespace',
      undefined,
    ],
  ] as const)(
    'rejects malformed worksheet structure %#',
    async (xml, message, cell) => {
      expect((await captureParseError(xml)).diagnostic).toEqual({
        ...(cell === undefined ? {} : { cell }),
        code: 'invalid-document-structure',
        message,
        part: PART,
        severity: 'error',
      });
    },
  );

  it.each([
    ['<row r="0"/>', 'Worksheet rows are out of order'],
    ['<row r="01"/>', 'Worksheet row reference is invalid'],
    ['<row r="+1"/>', 'Worksheet row reference is invalid'],
    ['<row r="1.0"/>', 'Worksheet row reference is invalid'],
    ['<row r="4294967296"/>', 'Worksheet row reference is invalid'],
    ['<row r="9007199254740992"/>', 'Worksheet row reference is invalid'],
    ['<row r="2"/><row r="2"/>', 'Worksheet rows are out of order'],
    ['<row r="2"/><row r="1"/>', 'Worksheet rows are out of order'],
    ['<row hidden="yes"/>', 'Worksheet row hidden flag is invalid'],
    ['<row outlineLevel="8"/>', 'Worksheet row outline level is invalid'],
    ['<row outlineLevel="01"/>', 'Worksheet row outline level is invalid'],
    ['<row ht="-1"/>', 'Worksheet row height is invalid'],
    ['<row ht="1e2"/>', 'Worksheet row height is invalid'],
    [
      `<row ht="${'9'.repeat(400)}"/>`,
      'Worksheet row height is outside the finite range',
    ],
  ] as const)('rejects invalid row value %#', async (rows, message) => {
    const error = await captureParseError(
      worksheet(`<sheetData>${rows}</sheetData>`),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-document-value',
      message,
      part: PART,
      severity: 'error',
    });
  });

  it.each([
    ['<c r="$A$1"/>', 'Worksheet cell reference is invalid', undefined],
    ['<c r="A0"/>', 'Worksheet cell reference is invalid', undefined],
    ['<c r="XFE1"/>', 'Worksheet cell reference is invalid', undefined],
    [
      '<c r="A2"/>',
      'Worksheet cell reference does not belong to its row',
      'A2',
    ],
    ['<c r="B1"/><c r="A1"/>', 'Worksheet cells are out of order', 'A1'],
    ['<c r="A1"/><c r="A1"/>', 'Worksheet cells are out of order', 'A1'],
    ['<c t="x"/>', 'Worksheet cell type is invalid', 'A1'],
    ['<c s="-1"/>', 'Worksheet cell style index is invalid', 'A1'],
    ['<c s="01"/>', 'Worksheet cell style index is invalid', 'A1'],
    ['<c s="4294967296"/>', 'Worksheet cell style index is invalid', 'A1'],
  ] as const)('rejects invalid cell value %#', async (cells, message, cell) => {
    const error = await captureParseError(
      worksheet(`<sheetData><row>${cells}</row></sheetData>`),
    );
    expect(error.diagnostic).toEqual({
      ...(cell === undefined ? {} : { cell }),
      code: 'invalid-document-value',
      message,
      part: PART,
      severity: 'error',
    });
  });

  it.each([
    ['<c><v>1</v><v>2</v></c>', 'Worksheet cell value structure is invalid'],
    [
      '<c t="inlineStr"><v>A</v></c>',
      'Worksheet cell value structure is invalid',
    ],
    [
      '<c><is><t>A</t></is></c>',
      'Worksheet inline-string structure is invalid',
    ],
    [
      '<c t="inlineStr"><is/><is/></c>',
      'Worksheet inline-string structure is invalid',
    ],
    [
      '<c t="inlineStr"><is><t>A</t><t>B</t></is></c>',
      'Worksheet inline-string plain text is out of order',
    ],
    [
      '<c t="inlineStr"><is><t>A</t><r><t>B</t></r></is></c>',
      'Worksheet inline-string rich text is out of order',
    ],
    [
      '<c t="inlineStr"><is><r><t>A</t></r><t>B</t></is></c>',
      'Worksheet inline-string plain text is out of order',
    ],
    [
      '<c t="inlineStr"><is><r/></is></c>',
      'Worksheet inline-string run text is missing',
    ],
    [
      '<c t="inlineStr"><is><r><t>A</t><t>B</t></r></is></c>',
      'Worksheet inline-string run has duplicate text',
    ],
    [
      '<c t="inlineStr"><is>text</is></c>',
      'Worksheet text is outside a value or inline-string text element',
    ],
  ] as const)('rejects invalid cell structure %#', async (cellXml, message) => {
    const error = await captureParseError(
      worksheet(`<sheetData><row>${cellXml}</row></sheetData>`),
    );
    expect(error.diagnostic).toEqual({
      cell: 'A1',
      code: 'invalid-document-structure',
      message,
      part: PART,
      severity: 'error',
    });
  });

  it.each(['invalid', 'preserved', ''])(
    'rejects inline xml:space value %#',
    async (space) => {
      const error = await captureParseError(
        worksheet(
          `<sheetData><row><c t="inlineStr"><is><t xml:space="${space}">A</t></is></c></row></sheetData>`,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        cell: 'A1',
        code: 'invalid-document-value',
        message: 'Inline-string xml:space value is invalid',
      });
    },
  );

  it('parses normal formulas with typed cached and missing results without evaluating them', async () => {
    await expect(
      parse(
        worksheet(`<sheetData><row>
          <c r="A1"><f>1+2</f><v>3</v></c>
          <c r="B1" t="str"><f>UNKNOWN(A1)</f><v>cached</v></c>
          <c r="C1" t="b"><f>TRUE()</f><v>1</v></c>
          <c r="D1" t="e"><f>1/0</f><v>#DIV/0!</v></c>
          <c r="E1" t="d"><f>DATE(2024,1,2)</f><v>2024-01-02</v></c>
          <c r="F1"><f>NOW()</f></c>
        </row></sheetData>`),
      ),
    ).resolves.toEqual({
      rows: [
        {
          cells: [
            {
              address: 'A1',
              column: 1,
              content: {
                cached: { kind: 'number', value: 3 },
                formula: { expression: '1+2', kind: 'normal' },
                kind: 'formula',
              },
            },
            {
              address: 'B1',
              column: 2,
              content: {
                cached: { kind: 'text', text: 'cached' },
                formula: { expression: 'UNKNOWN(A1)', kind: 'normal' },
                kind: 'formula',
              },
            },
            {
              address: 'C1',
              column: 3,
              content: {
                cached: { kind: 'boolean', value: true },
                formula: { expression: 'TRUE()', kind: 'normal' },
                kind: 'formula',
              },
            },
            {
              address: 'D1',
              column: 4,
              content: {
                cached: { code: '#DIV/0!', kind: 'error' },
                formula: { expression: '1/0', kind: 'normal' },
                kind: 'formula',
              },
            },
            {
              address: 'E1',
              column: 5,
              content: {
                cached: {
                  kind: 'date',
                  normalized: '2024-01-02',
                  precision: 'date',
                  source: { kind: 'iso', value: '2024-01-02' },
                },
                formula: { expression: 'DATE(2024,1,2)', kind: 'normal' },
                kind: 'formula',
              },
            },
            {
              address: 'F1',
              column: 6,
              content: {
                cached: { kind: 'missing' },
                formula: { expression: 'NOW()', kind: 'normal' },
                kind: 'formula',
              },
            },
          ],
          index: 1,
        },
      ],
    });
  });

  it('expands shared formulas while retaining a master outside selection', async () => {
    const selection: XlsxResolvedSheetSelection = {
      endRowPrefix: [2],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 2, row: 2 },
          reference: 'B2',
          start: { column: 2, row: 2 },
        },
      ],
    };
    const budget = createXlsxWorksheetBudget(EMPTY_STRINGS);
    const result = await parse(
      worksheet(`<sheetData>
        <row r="1">
          <c r="A1"><f t="shared" si="7" ref="A1:B2">A1+$B$1+"A1"</f><v>1</v></c>
          <c r="B1"><f t="shared" si="7"/><v>2</v></c>
        </row>
        <row r="2">
          <c r="A2"><f t="shared" si="7"/><v>3</v></c>
          <c r="B2"><f t="shared" si="7"/><v>4</v></c>
        </row>
      </sheetData>`),
      { budget, selection },
    );

    expect(result).toEqual({
      rows: [
        {
          cells: [
            {
              address: 'B2',
              column: 2,
              content: {
                cached: { kind: 'number', value: 4 },
                formula: {
                  expression: 'B2+$B$1+"A1"',
                  kind: 'normal',
                },
                kind: 'formula',
              },
            },
          ],
          index: 2,
        },
      ],
    });
    expect(budget.formulaGroups).toBe(1);
    expect(budget.formulaCharacters).toBe(48);
  });

  it('parses array and data-table formula groups with explicit ranges', async () => {
    const result = await parse(
      worksheet(`<sheetData><row>
        <c r="A1"><f t="array" ref="A1:B2">SUM(C1:C2)</f><v>3</v></c>
        <c r="C1"><f t="dataTable" ref="C1:D2"/></c>
      </row></sheetData>`),
    );

    expect(result.rows[0]!.cells.map((cell) => cell.content)).toEqual([
      {
        cached: { kind: 'number', value: 3 },
        formula: {
          expression: 'SUM(C1:C2)',
          kind: 'array',
          range: {
            end: { column: 2, row: 2 },
            reference: 'A1:B2',
            start: { column: 1, row: 1 },
          },
        },
        kind: 'formula',
      },
      {
        cached: { kind: 'missing' },
        formula: {
          expression: '',
          kind: 'data-table',
          range: {
            end: { column: 4, row: 2 },
            reference: 'C1:D2',
            start: { column: 3, row: 1 },
          },
        },
        kind: 'formula',
      },
    ]);
  });

  it.each([
    ['<f t="future">A1</f>', 'Formula type is invalid'],
    ['<f/>', 'Normal formula is invalid'],
    ['<f ref="A1">A1</f>', 'Normal formula is invalid'],
    ['<f si="1">A1</f>', 'Normal formula is invalid'],
    ['<f>=A1</f>', 'Formula expression must not include a leading equals sign'],
    ['<f t="array" ref="A1"/>', 'Array formula is empty'],
    ['<f t="array">A1</f>', 'Formula range is invalid'],
    [
      '<f t="array" ref="B1">A1</f>',
      'Grouped formula must start at its owning cell',
    ],
    [
      '<f t="array" ref="A2">A1</f>',
      'Grouped formula must start at its owning cell',
    ],
    [
      '<f t="array" ref="A1" si="1">A1</f>',
      'Grouped formula shared index is invalid',
    ],
    ['<f t="dataTable"/>', 'Formula range is invalid'],
    ['<f t="shared">A1</f>', 'Shared formula index is invalid'],
    ['<f t="shared" si="-1">A1</f>', 'Shared formula index is invalid'],
    ['<f t="shared" si="4294967296">A1</f>', 'Shared formula index is invalid'],
    ['<f t="shared" si="1" ref="A1"/>', 'Shared formula master is invalid'],
    [
      '<f t="shared" si="1" ref="B1">A1</f>',
      'Shared formula master must start at its owning cell',
    ],
    [
      '<f t="shared" si="1" ref="A2">A1</f>',
      'Shared formula master must start at its owning cell',
    ],
  ])('rejects invalid formula %#', async (formula, message) => {
    const error = await captureParseError(
      worksheet(`<sheetData><row><c>${formula}</c></row></sheetData>`),
    );
    expect(error.diagnostic).toEqual({
      cell: 'A1',
      code: 'invalid-formula',
      message,
      part: PART,
      severity: 'error',
    });
  });

  it.each([
    [
      '<c><f t="shared" si="1" ref="A1:B1">A1</f></c><c><f t="shared" si="1" ref="B1">B1</f></c>',
      'Shared formula master is invalid',
      'B1',
    ],
    [
      '<c><f t="shared" si="1">A1</f></c>',
      'Shared formula dependent contains an expression',
      'A1',
    ],
    [
      '<c><f t="shared" si="1"/></c>',
      'Shared formula master is missing or does not own the cell',
      'A1',
    ],
    [
      '<c><f t="shared" si="1" ref="A1:A1">A1</f></c><c><f t="shared" si="1"/></c>',
      'Shared formula master is missing or does not own the cell',
      'B1',
    ],
    [
      '<c><f t="shared" si="1" ref="A1:A1">A1</f></c></row><row><c><f t="shared" si="1"/></c>',
      'Shared formula master is missing or does not own the cell',
      'A2',
    ],
    [
      '<c><f t="shared" si="1" ref="A1:B1">XFD1</f></c><c><f t="shared" si="1"/></c>',
      'Shared formula translation is outside the worksheet grid',
      'B1',
    ],
  ])(
    'rejects invalid shared formula group %#',
    async (cells, message, cell) => {
      const error = await captureParseError(
        worksheet(`<sheetData><row>${cells}</row></sheetData>`),
      );
      expect(error.diagnostic).toMatchObject({
        cell,
        code: 'invalid-formula',
        message,
      });
    },
  );

  it.each([
    ['<c><v>1</v><f>A1</f></c>', 'Worksheet formula structure is invalid'],
    ['<c><f>A1</f><f>B1</f></c>', 'Worksheet formula structure is invalid'],
    [
      '<c t="inlineStr"><f>A1</f></c>',
      'Worksheet formula structure is invalid',
    ],
    ['<c t="s"><f>A1</f></c>', 'Worksheet formula structure is invalid'],
    ['<c><f><v>A1</v></f></c>', 'Worksheet element nesting is invalid'],
  ])('rejects invalid formula cell structure %#', async (cellXml, message) => {
    const error = await captureParseError(
      worksheet(`<sheetData><row>${cellXml}</row></sheetData>`),
    );
    expect(error.diagnostic).toMatchObject({
      cell: 'A1',
      code: 'invalid-document-structure',
      message,
    });
  });

  it('enforces per-formula, aggregate-formula, and group limits at exact boundaries', async () => {
    const formulas = worksheet(
      '<sheetData><row><c><f>AB</f></c><c><f>CD</f></c></row></sheetData>',
    );
    await expect(
      parse(formulas, {
        limits: { maxFormulaCharacters: 2, maxTotalFormulaCharacters: 4 },
      }),
    ).resolves.toMatchObject({ rows: [{ cells: [{}, {}] }] });
    await expect(
      parse(formulas, {
        limits: { maxFormulaCharacters: 1, maxTotalFormulaCharacters: 4 },
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxFormulaCharacters',
    });
    await expect(
      parse(formulas, {
        limits: { maxFormulaCharacters: 2, maxTotalFormulaCharacters: 3 },
      }),
    ).rejects.toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTotalFormulaCharacters',
    });

    const groups = worksheet(
      '<sheetData><row><c><f t="array" ref="A1">A</f></c><c><f t="array" ref="B1">B</f></c></row></sheetData>',
    );
    await expect(
      parse(groups, { limits: { maxFormulaGroups: 2 } }),
    ).resolves.toMatchObject({ rows: [{ cells: [{}, {}] }] });
    await expect(
      parse(groups, { limits: { maxFormulaGroups: 1 } }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxFormulaGroups',
    });
  });

  it('accepts the maximum shared formula index', async () => {
    await expect(
      parse(
        worksheet(
          '<sheetData><row><c><f t="shared" si="4294967295" ref="A1">A1</f></c></row></sheetData>',
        ),
      ),
    ).resolves.toMatchObject({
      rows: [
        {
          cells: [
            {
              content: {
                formula: { expression: 'A1', kind: 'normal' },
              },
            },
          ],
        },
      ],
    });
  });

  it('charges selected formula cached text to the returned text budget', async () => {
    const xml = worksheet(
      '<sheetData><row><c t="str"><f>TEXT(A1)</f><v>AB</v></c></row></sheetData>',
    );
    await expect(
      parse(xml, { limits: { maxTextCharacters: 2 } }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    await expect(
      parse(xml, { limits: { maxTextCharacters: 1 } }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxTextCharacters',
    });
  });

  it.each([
    ['maxRowsPerWorksheet', '<row r="2"/>', 2],
    ['maxColumnsPerWorksheet', '<row><c r="B1"/></row>', 2],
    ['maxScannedCells', '<row><c/><c/></row>', 2],
    ['maxReturnedCells', '<row><c/><c/></row>', 2],
  ] as const)('enforces %s at one over', async (limitName, rowXml, actual) => {
    const overrides = {
      [limitName]: 1,
      ...(limitName === 'maxReturnedCells' ? { maxScannedCells: 2 } : {}),
    };
    await expect(
      parse(worksheet(`<sheetData>${rowXml}</sheetData>`), {
        limits: overrides,
      }),
    ).rejects.toMatchObject({
      actual,
      limit: 1,
      limitName,
      name: 'XlsxResourceLimitError',
      part: PART,
    } satisfies Partial<XlsxResourceLimitError>);
  });

  it('accepts row, column, scanned, and returned counts exactly at limits', async () => {
    const budget = createXlsxWorksheetBudget(EMPTY_STRINGS);
    await expect(
      parse(
        worksheet(
          '<sheetData><row r="2"><c r="B2"/><c r="C2"/></row></sheetData>',
        ),
        {
          budget,
          limits: {
            maxColumnsPerWorksheet: 3,
            maxReturnedCells: 2,
            maxRowsPerWorksheet: 2,
            maxScannedCells: 2,
          },
        },
      ),
    ).resolves.toMatchObject({ rows: [{ index: 2 }] });
    expect(budget.returnedCells).toBe(2);
    expect(budget.scannedCells).toBe(2);
  });

  it('enforces omitted column references at the configured boundary', async () => {
    await expect(
      parse(worksheet('<sheetData><row><c/></row></sheetData>'), {
        limits: { maxColumnsPerWorksheet: 1 },
      }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    await expect(
      parse(worksheet('<sheetData><row><c/><c/></row></sheetData>'), {
        limits: { maxColumnsPerWorksheet: 1 },
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxColumnsPerWorksheet',
      name: 'XlsxResourceLimitError',
      part: PART,
    } satisfies Partial<XlsxResourceLimitError>);
  });

  it('enforces expanded text characters including shared-string copies', async () => {
    const strings: XlsxSharedStringTable = {
      part: 'strings.xml',
      values: [{ text: 'AB' }],
    };
    const xml = worksheet(
      '<sheetData><row><c t="s"><v>0</v></c></row></sheetData>',
    );
    await expect(
      parse(xml, { limits: { maxTextCharacters: 4 }, strings }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    await expect(
      parse(xml, { limits: { maxTextCharacters: 3 }, strings }),
    ).rejects.toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTextCharacters',
      name: 'XlsxResourceLimitError',
      part: PART,
    } satisfies Partial<XlsxResourceLimitError>);
  });

  it('enforces rich-run objects for shared copies and inline output', async () => {
    const strings: XlsxSharedStringTable = {
      part: 'strings.xml',
      values: [{ runs: [{ text: 'A' }], text: 'A' }],
    };
    const sharedXml = worksheet(
      '<sheetData><row><c t="s"><v>0</v></c></row></sheetData>',
    );
    await expect(
      parse(sharedXml, {
        limits: { maxRichTextRuns: 2, maxTextCharacters: 2 },
        strings,
      }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    await expect(
      parse(sharedXml, {
        limits: { maxRichTextRuns: 1, maxTextCharacters: 2 },
        strings,
      }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxRichTextRuns',
      name: 'XlsxResourceLimitError',
    } satisfies Partial<XlsxResourceLimitError>);

    const inlineXml = worksheet(
      '<sheetData><row><c t="inlineStr"><is><r><t>A</t></r><r><t>B</t></r></is></c></row></sheetData>',
    );
    await expect(
      parse(inlineXml, { limits: { maxRichTextRuns: 2 } }),
    ).resolves.toMatchObject({ rows: [{ cells: [{ address: 'A1' }] }] });
    await expect(
      parse(inlineXml, { limits: { maxRichTextRuns: 1 } }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxRichTextRuns',
      name: 'XlsxResourceLimitError',
    } satisfies Partial<XlsxResourceLimitError>);
  });

  it('requires the worksheet part through the bounded reader', async () => {
    const zip = new JSZip();
    const resolved = limits();
    const reader = new XlsxPartReader(zip, [], resolved);
    await expect(
      parseXlsxWorksheetPart(
        PART,
        'transitional',
        reader,
        resolved,
        EMPTY_STRINGS,
        createXlsxWorksheetBudget(EMPTY_STRINGS),
      ),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        message: `Required XLSX part is missing: ${PART}`,
        part: PART,
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });
});
