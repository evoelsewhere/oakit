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
