import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import { patchXlsxTableStructure } from '../../src/formats/xlsx/roundtrip/table-structure-patch';
import { patchXlsxWorksheetStructure } from '../../src/formats/xlsx/roundtrip/worksheet-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
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
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><selection activeCell="B2" activeCellId="1" sqref="A1:A2 B2"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData><protectedRanges><protectedRange name="Input" sqref="A1:B2"/></protectedRanges><autoFilter ref="A1:B2"><sortState ref="A1:B2"><sortCondition ref="A1:A2"/></sortState></autoFilter><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><conditionalFormatting sqref="A1:B2"><cfRule type="top10" priority="1" rank="1"/></conditionalFormatting><dataValidations count="1" disablePrompts="1"><dataValidation sqref="A1:B2"/></dataValidations><hyperlinks><hyperlink ref="A1:B1" location="Sheet1!A1"/></hyperlinks><rowBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="1" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="0"><brk id="2" min="0" max="1" pt="1"/></colBreaks></worksheet>`,
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
    expect(sheet.hyperlinks.map((link) => link.range.reference)).toEqual([
      'A2:B2',
    ]);
    expect(sheet.autoFilter?.range.reference).toBe('A2:B3');
    expect(sheet.autoFilter?.sort?.range.reference).toBe('A2:B3');
    expect(sheet.autoFilter?.sort?.conditions[0]?.range.reference).toBe(
      'A2:A3',
    );
    expect(sheet.dataValidations[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.dataValidationSettings).toEqual({ disablePrompts: true });
    expect(sheet.conditionalFormattings[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.protectedRanges[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.print?.rowBreaks?.[0]).toMatchObject({
      end: 1,
      position: 3,
      start: 0,
    });
    expect(sheet.print?.columnBreaks?.[0]).toMatchObject({
      end: 2,
      position: 2,
      start: 1,
    });
    expect(sheet.views[0]?.topLeftCell).toBe('A3');
    expect(sheet.views[0]?.selections[0]).toMatchObject({
      activeCell: 'B3',
      activeCellId: 1,
      ranges: [{ reference: 'A2:A3' }, { reference: 'B3' }],
    });
    expect(result.report.level).toBe('R2');
  });

  it('keeps worksheet tables and their owned parts aligned', async () => {
    const tablePart = 'xl/tables/table1.xml';
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/${tablePart}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`;
    const generatedSource = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${XLSX_OFFICE_REL_TYPE}officeDocument" Target="xl/workbook.xml"/><Relationship Id="unowned-table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="xl/tables/table2.xml"/></Relationships>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="../tables/table1.xml"/><Relationship Id="external-table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="https://example.invalid/table.xml" TargetMode="External"/><Relationship Id="internal-link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="../tables/table2.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1" t="str"><v>A</v></c><c r="B1" t="str"><v>B</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row></sheetData><tableParts count="1"><tablePart r:id="table"/></tableParts></worksheet>`,
      [tablePart]: `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Table1" displayName="Table1" ref="A1:B3"><autoFilter ref="A1:B3"><sortState ref="A1:B3"><sortCondition ref="A2:A3"/></sortState></autoFilter><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`,
      'xl/tables/table2.xml': `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Unused" displayName="Unused" ref="D10:E12"><tableColumns count="2"><tableColumn id="1" name="D"/><tableColumn id="2" name="E"/></tableColumns></table>`,
    });
    const datedSource = await JSZip.loadAsync(generatedSource);
    const tableBytes = await datedSource.file(tablePart)!.async('uint8array');
    datedSource.file(tablePart, tableBytes, {
      date: new Date('2001-02-03T04:05:06.000Z'),
    });
    const source = await datedSource.generateAsync({ type: 'uint8array' });
    const snapshot = await readXlsxRoundTrip(source);
    const structuralOperation = {
      count: 1,
      index: 2,
      kind: 'insert-rows' as const,
      operationId: 'insert-table-row',
      sheetKey: snapshot.document.sheets[0]!.key,
    };
    const edited = await applyXlsxEdits(snapshot, [structuralOperation]);
    const first = await writeXlsxRoundTrip(edited);
    const second = await writeXlsxRoundTrip(portable(edited));
    expect(second.data).toEqual(first.data);
    expect(
      first.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/tables/table1.xml', 'xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(first.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.tables[0]?.range.reference).toBe('A1:B4');
    expect(sheet.tables[0]?.autoFilter?.range.reference).toBe('A1:B4');
    expect(
      sheet.tables[0]?.autoFilter?.sort?.conditions[0]?.range.reference,
    ).toBe('A3:A4');
    expect(first.report.level).toBe('R2');

    const sourceZip = await JSZip.loadAsync(source);
    const outputZip = await JSZip.loadAsync(first.data);
    expect(outputZip.file(tablePart)!.date).toEqual(
      sourceZip.file(tablePart)!.date,
    );
    const request = {
      count: structuralOperation.count,
      index: structuralOperation.index,
      kind: structuralOperation.kind,
      operationId: structuralOperation.operationId,
    };
    const worksheetPatch = patchXlsxWorksheetStructure(
      await sourceZip.file('xl/worksheets/sheet1.xml')!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const tablePatch = patchXlsxTableStructure(
      await sourceZip.file(tablePart)!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      tablePart,
    );
    const patchBytes = worksheetPatch.patchBytes + tablePatch.patchBytes;
    const patchCount = worksheetPatch.patchCount + tablePatch.patchCount;
    const generatedXmlBytes = first.report.parts
      .filter((part) => part.disposition === 'patch')
      .reduce((total, part) => total + part.byteLength, 0);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxDependencyEdges: 2,
          maxDirtyParts: 2,
          maxGeneratedXmlBytes: generatedXmlBytes,
          maxPatchBytes: patchBytes,
          maxPatchCount: patchCount,
          maxPatchedParts: 2,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxDependencyEdges', 1],
      ['maxDirtyParts', 1],
      ['maxGeneratedXmlBytes', generatedXmlBytes - 1],
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
      ['maxPatchedParts', 1],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(edited, { limits: { [limitName]: limit } }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }

    const rowEdit = await applyXlsxEdits(snapshot, [
      {
        hidden: true,
        kind: 'set-row',
        operationId: 'table-row-property',
        row: 1,
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    await expect(writeXlsxRoundTrip(rowEdit)).rejects.toMatchObject({
      diagnostic: { featureClass: 'unsupported-part', part: tablePart },
    });
    const outside = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 5,
        kind: 'insert-rows',
        operationId: 'outside-table',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const outsideResult = await writeXlsxRoundTrip(outside);
    expect(
      outsideResult.report.parts.find((part) => part.name === tablePart)
        ?.disposition,
    ).toBe('copy');
    for (const operation of [
      {
        count: 1,
        index: 2,
        kind: 'delete-rows' as const,
        operationId: 'delete-table-data-row',
      },
      {
        count: 1,
        index: 3,
        kind: 'insert-columns' as const,
        operationId: 'insert-outside-table-column',
      },
      {
        count: 1,
        index: 3,
        kind: 'delete-columns' as const,
        operationId: 'delete-outside-table-column',
      },
    ]) {
      const candidate = await applyXlsxEdits(snapshot, [
        {
          ...operation,
          sheetKey: snapshot.document.sheets[0]!.key,
        },
      ]);
      await expect(writeXlsxRoundTrip(candidate)).resolves.toMatchObject({
        report: { level: 'R2' },
      });
    }
  });
});
