import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { patchXlsxWorksheetStructure } from '../../src/formats/xlsx/roundtrip/worksheet-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { XLSX_SPREADSHEET_NS } from '../black-box/xlsx-package';

const PART = 'xl/worksheets/sheet1.xml';

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected structural patch to fail');
}

function source(): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3" hidden="1" spans="3:3"><c r="C3" s="2"><v>3</v></c></row></sheetData></worksheet>`;
}

describe('XLSX worksheet structural patching', () => {
  it('applies ordered row and column insertions and deletions', () => {
    const result = patchXlsxWorksheetStructure(
      bytes(source()),
      [
        { count: 2, index: 2, kind: 'insert-rows', operationId: 'insert-rows' },
        { count: 1, index: 1, kind: 'delete-rows', operationId: 'delete-rows' },
        {
          count: 1,
          index: 2,
          kind: 'insert-columns',
          operationId: 'insert-columns',
        },
        {
          count: 1,
          index: 1,
          kind: 'delete-columns',
          operationId: 'delete-columns',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="4" hidden="1"><c r="C4" s="2"><v>3</v></c></row></sheetData></worksheet>`,
    );
    expect(result.patchCount).toBe(8);
  });

  it('preserves an owned byte copy when no operation is requested', () => {
    const input = bytes(source());
    const result = patchXlsxWorksheetStructure(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result).toEqual({ data: input, patchBytes: 0, patchCount: 0 });
    expect(result.data).not.toBe(input);
  });

  it('distinguishes every insertion and deletion boundary', () => {
    const rows = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>${[
      1, 2, 3, 4, 5,
    ]
      .map((row) => `<row r="${row}"><c r="A${row}"><v>${row}</v></c></row>`)
      .join('')}</sheetData></worksheet>`;
    const deletedRows = patchXlsxWorksheetStructure(
      bytes(rows),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'rows' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deletedRows.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>4</v></c></row><row r="3"><c r="A3"><v>5</v></c></row></sheetData></worksheet>`,
    );
    const insertedRows = patchXlsxWorksheetStructure(
      bytes(rows),
      [{ count: 1, index: 2, kind: 'insert-rows', operationId: 'rows' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(insertedRows.data)).toContain(
      '<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c></row>',
    );

    const columns = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1" spans="1:5"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c><c r="E1"><v>5</v></c></row></sheetData></worksheet>`;
    const deletedColumns = patchXlsxWorksheetStructure(
      bytes(columns),
      [{ count: 2, index: 2, kind: 'delete-columns', operationId: 'columns' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deletedColumns.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>4</v></c><c r="C1"><v>5</v></c></row></sheetData></worksheet>`,
    );
    expect(deletedColumns.patchCount).toBe(5);
    const insertedColumns = patchXlsxWorksheetStructure(
      bytes(columns),
      [{ count: 1, index: 2, kind: 'insert-columns', operationId: 'columns' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(insertedColumns.data)).toContain(
      '<c r="A1"><v>1</v></c><c r="C1"><v>2</v></c>',
    );
    expect(insertedColumns.patchCount).toBe(5);
  });

  it('patches only direct prefixed rows and cells', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><wrapper><s:row r="6"/><s:sheetData><s:row r="9"/></s:sheetData></wrapper><s:sheetData><other r="1"/><wrapper><s:row r="7"/></wrapper><s:row r="1"><wrapper><s:c r="Z9"/></wrapper><other r="A1"/><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData><wrapper><s:row r="8"/></wrapper></s:worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'prefixed' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><wrapper><s:row r="6"/><s:sheetData><s:row r="9"/></s:sheetData></wrapper><s:sheetData><other r="1"/><wrapper><s:row r="7"/></wrapper><s:row r="2"><wrapper><s:c r="Z9"/></wrapper><other r="A1"/><s:c r="A2"><s:v>1</s:v></s:c></s:row></s:sheetData><wrapper><s:row r="8"/></wrapper></s:worksheet>`,
    );
  });

  it('transforms declared dimensions and merged ranges with exact counts', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:C5"/><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells count="2"><mergeCell ref="A2:B3"/><mergeCell ref="C1:C5"/></mergeCells><hyperlinks><hyperlink ref="A2:B3" location="Removed!A1"/><hyperlink ref="C1:C5" location="Kept!A1"/></hyperlinks></worksheet>`;
    const deleted = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'layout' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deleted.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells count="1"><mergeCell ref="C1:C3"/></mergeCells><hyperlinks><hyperlink ref="C1:C3" location="Kept!A1"/></hyperlinks></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A2:B3"/><sheetData/><mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells><hyperlinks><hyperlink ref="A2:B3" location="Removed!A1"/></hyperlinks></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-layout',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
    const removedSort = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C5"><sortState ref="A2:C3"><sortCondition ref="A2:A3"/></sortState></autoFilter></worksheet>`,
      ),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'sort' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedSort.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C3"></autoFilter></worksheet>`,
    );
  });

  it('selects only direct filter and sort nodes and avoids no-op patches', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:autoFilter ref="Z9"><s:sortState ref="Z9"><s:sortCondition ref="Z9"/></s:sortState></s:autoFilter><s:sortState ref="Z8"/></wrapper><x:autoFilter ref="Z9"/><s:autoFilter ref="A1:C3"><wrapper><s:sortState ref="Z9"/></wrapper><wrapper><s:sortCondition ref="Z8"/></wrapper><x:sortState ref="Z9"/><s:sortState ref="A1:C3"><other ref="Z9"/><wrapper><s:sortCondition ref="Z9"/></wrapper><x:sortCondition ref="Z9"/><s:sortCondition ref="A1:A3"/></s:sortState><wrapper><s:sortCondition ref="Z9"/></wrapper></s:autoFilter><wrapper><s:sortState ref="Z9"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'unchanged-filter',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'changed-filter',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:autoFilter ref="A2:C4">');
    expect(output).toContain('<s:sortState ref="A2:C4">');
    expect(output).toContain('<s:sortCondition ref="A2:A4"/>');
    expect(output).toContain(
      '<wrapper><s:autoFilter ref="Z9"><s:sortState ref="Z9"><s:sortCondition ref="Z9"/></s:sortState></s:autoFilter><s:sortState ref="Z8"/></wrapper>',
    );
    expect(output).toContain('<wrapper><s:sortState ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z8"/></wrapper>');
  });

  it('does not treat a following worksheet sort state as part of an auto-filter', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><s:sheetData/><s:autoFilter ref="A1:C3"/><wrapper><s:sortState ref="Z7"/></wrapper></s:worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 0,
          kind: 'insert-rows',
          operationId: 'filter-ownership',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );

    expect(new TextDecoder().decode(result.data)).toBe(
      `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><s:sheetData/><s:autoFilter ref="A2:C4"/><wrapper><s:sortState ref="Z7"/></wrapper></s:worksheet>`,
    );
  });

  it('transforms worksheet filter and sort ranges', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C5"><filterColumn colId="0"/><sortState ref="A1:C5"><sortCondition ref="A2:A3"/><sortCondition ref="B1:B5" descending="1"/></sortState></autoFilter></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'filter' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C3"><filterColumn colId="0"/><sortState ref="A1:C3"><sortCondition ref="B1:B3" descending="1"/></sortState></autoFilter></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A2:C3"/></worksheet>`,
      ),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'filter' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
  });

  it('selects only direct layout nodes and avoids no-op patches', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><wrapper><s:dimension ref="Z9"/></wrapper><x:dimension ref="Z9"/><s:dimension ref="A1:B2"/><s:sheetData/><wrapper><s:hyperlinks><s:hyperlink ref="Z9" location="Nested!A1"/></s:hyperlinks><s:hyperlink ref="Z8"/></wrapper><x:hyperlinks><x:hyperlink ref="Z9"/></x:hyperlinks><s:hyperlinks><other ref="Z9"/><wrapper><s:hyperlink ref="Z9"/></wrapper><x:hyperlink ref="Z9"/><s:hyperlink ref="A1:B2" location="Real!A1"/></s:hyperlinks><wrapper><s:hyperlink ref="Z9"/></wrapper><wrapper><s:mergeCells count="1"><s:mergeCell ref="Z9"/></s:mergeCells><s:mergeCell ref="Z9"/></wrapper><s:mergeCells count="1"><other ref="Z9"/><wrapper><s:mergeCell ref="Z9"/></wrapper><x:mergeCell ref="Z9"/><s:mergeCell ref="A1:B2"/></s:mergeCells><wrapper><s:mergeCell ref="Z9"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 5, kind: 'insert-rows', operationId: 'unchanged' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'changed' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:dimension ref="A2:B3"/>');
    expect(output).toContain('<s:mergeCell ref="A2:B3"/>');
    expect(output).toContain('<s:hyperlink ref="A2:B3" location="Real!A1"/>');
    expect(output).toContain(
      '<wrapper><s:hyperlinks><s:hyperlink ref="Z9" location="Nested!A1"/></s:hyperlinks><s:hyperlink ref="Z8"/></wrapper>',
    );
    expect(output).toContain('<x:hyperlink ref="Z9"/>');
    expect(output).toContain('<wrapper><s:dimension ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:mergeCell ref="Z9"/></wrapper>');
    expect(output).toContain(
      '<wrapper><s:mergeCells count="1"><s:mergeCell ref="Z9"/></s:mergeCells><s:mergeCell ref="Z9"/></wrapper>',
    );
    expect(output).toContain('<x:mergeCell ref="Z9"/>');
  });

  it('rejects malformed dimension and merge references exactly', () => {
    for (const [xml, message] of [
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="bad"/><sheetData/></worksheet>`,
        'XLSX structural dimension reference is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><mergeCells><mergeCell ref="bad"/></mergeCells></worksheet>`,
        'XLSX structural merged range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><hyperlinks><hyperlink ref="bad"/></hyperlinks></worksheet>`,
        'XLSX structural hyperlink range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="bad"/></worksheet>`,
        'XLSX structural auto-filter range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1"><sortState ref="bad"/></autoFilter></worksheet>`,
        'XLSX structural sort range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1"><sortState ref="A1"><sortCondition ref="bad"/></sortState></autoFilter></worksheet>`,
        'XLSX structural sort-condition range is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(xml),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-layout',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }
  });

  it('rejects unsafe roots, missing sheetData, and malformed references', () => {
    for (const [xml, message] of [
      [
        `<outer>${source()}</outer>`,
        'XLSX worksheet root cannot patch structure',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"/>`,
        'XLSX worksheet sheetData cannot patch structure',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row><c r="A1"/></row></sheetData></worksheet>`,
        'XLSX structural target row reference is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c/></row></sheetData></worksheet>`,
        'XLSX structural target cell reference is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(xml),
            [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'bad' }],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ message, part: PART });
    }
    expect(
      capture(() =>
        patchXlsxWorksheetStructure(
          bytes(
            `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row/></sheetData></worksheet>`,
          ),
          [
            {
              count: 2,
              index: 3,
              kind: 'insert-rows',
              operationId: 'bad-range',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'worksheet-structure-xml',
      operationId: 'bad-range',
      range: '3:4',
    });
  });

  it('enforces patch count, patch bytes, and generated bytes exactly', () => {
    const request = [
      {
        count: 1,
        index: 1,
        kind: 'insert-rows' as const,
        operationId: 'insert',
      },
    ];
    const successful = patchXlsxWorksheetStructure(
      bytes(source()),
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(successful.patchBytes).toBe(33);
    expect(() =>
      patchXlsxWorksheetStructure(
        bytes(source()),
        request,
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
          maxPatchBytes: successful.patchBytes,
          maxPatchCount: successful.patchCount,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', successful.data.byteLength - 1],
      ['maxPatchBytes', successful.patchBytes - 1],
      ['maxPatchCount', successful.patchCount - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(source()),
            request,
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limit, limitName, part: PART });
    }

    const first = patchXlsxWorksheetStructure(
      bytes(source()),
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    const secondRequest = [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows' as const,
        operationId: 'second',
      },
    ];
    const second = patchXlsxWorksheetStructure(
      first.data,
      secondRequest,
      defaultXlsxWriteLimits(),
      PART,
    );
    const aggregateBytes = first.patchBytes + second.patchBytes;
    const aggregateCount = first.patchCount + second.patchCount;
    const individualBytes = Math.max(first.patchBytes, second.patchBytes);
    const individualCount = Math.max(first.patchCount, second.patchCount);
    for (const [limitName, limit, actual] of [
      ['maxPatchBytes', individualBytes, aggregateBytes],
      ['maxPatchCount', individualCount, aggregateCount],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(source()),
            [...request, ...secondRequest],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ actual, limit, limitName, part: PART });
    }
  });
});
