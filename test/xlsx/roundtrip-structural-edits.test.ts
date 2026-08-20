import { describe, expect, it } from 'vitest';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function portable<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

describe('XLSX verified structural row and column edits', () => {
  it('inserts and deletes authored cells in ordered atomic batches', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3" hidden="1"><c r="C3" s="0"><v>3</v></c></row></sheetData></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const sheetKey = snapshot.document.sheets[0]!.key;
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 2,
        index: 2,
        kind: 'insert-rows',
        operationId: 'insert-rows',
        sheetKey,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-rows',
        operationId: 'delete-rows',
        sheetKey,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns',
        operationId: 'insert-columns',
        sheetKey,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-columns',
        operationId: 'delete-columns',
        sheetKey,
      },
    ]);
    const first = await writeXlsxRoundTrip(portable(edited));
    const second = await writeXlsxRoundTrip(portable(edited));
    expect(second).toEqual(first);
    expect(first.report.level).toBe('R2');
    expect(
      first.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(first.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows).toEqual([
      expect.objectContaining({
        cells: [
          expect.objectContaining({
            address: 'C4',
            column: 3,
            content: { kind: 'value', value: { kind: 'number', value: 3 } },
            style: 0,
          }),
        ],
        hidden: true,
        index: 4,
      }),
    ]);
  });

  it('shifts a prefixed Strict worksheet without producer software', async () => {
    const strictSheetNs = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelNs =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelNs}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelNs}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheetNs}" xmlns:r="${strictRelNs}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 1,
        kind: 'insert-rows',
        operationId: 'strict-insert-row',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(result.report.level).toBe('R2');
    expect((await readXlsxRoundTrip(result.data)).source.conformance).toBe(
      'strict',
    );
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]).toMatchObject({
      cells: [expect.objectContaining({ address: 'A2' })],
      index: 2,
    });
  });

  it('keeps declared dimensions and merged ranges aligned', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 1,
        kind: 'insert-rows',
        operationId: 'insert-layout-row',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.declaredDimension?.reference).toBe('A2:B3');
    expect(sheet.mergedRanges.map((range) => range.reference)).toEqual([
      'A2:B2',
    ]);
    expect(result.report.level).toBe('R2');
  });
});
