import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  XlsxParseError,
} from '../../src/formats/xlsx';
import { createIndependentXlsx } from '../black-box/xlsx-package';

function arrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}

describe('public XLSX parser', () => {
  it('returns deterministic portable workbook metadata', async () => {
    const bytes = await createIndependentXlsx();
    const result = await parseXlsxWithDiagnostics(bytes);

    expect(result).toEqual({
      diagnostics: [],
      document: {
        sheets: [
          {
            columns: [],
            drawings: [],
            index: 0,
            kind: 'worksheet',
            mergedRanges: [],
            name: 'Sheet1',
            payload: 'full-sheet',
            rows: [],
            state: 'visible',
            tables: [],
          },
        ],
        styles: [],
        workbook: {
          calculation: {
            forceFullCalculation: false,
            fullCalculationOnLoad: false,
            mode: 'automatic',
          },
          dateSystem: '1900',
          definedNames: [],
        },
      },
    });
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('treats ArrayBuffer, Uint8Array, subarray, and Blob inputs identically', async () => {
    const bytes = await createIndependentXlsx();
    const padded = new Uint8Array(bytes.byteLength + 2);
    padded.set(bytes, 1);
    const inputs = [
      bytes,
      padded.subarray(1, padded.byteLength - 1),
      arrayBuffer(bytes),
      new Blob([arrayBuffer(bytes)]),
    ];

    const outputs = await Promise.all(inputs.map((input) => parseXlsx(input)));
    for (const output of outputs) expect(output).toEqual(outputs[0]);
  });

  it('does not mutate caller bytes or options', async () => {
    const bytes = await createIndependentXlsx();
    const before = bytes.slice();
    const limits = { maxInputBytes: bytes.byteLength };
    const options = { limits } as const;

    await parseXlsx(bytes, options);

    expect(bytes).toEqual(before);
    expect(options).toEqual({ limits: { maxInputBytes: bytes.byteLength } });
  });

  it('isolates concurrent parses and returns literal-equal results', async () => {
    const bytes = await createIndependentXlsx();
    const outputs = await Promise.all(
      Array.from({ length: 12 }, () => parseXlsx(bytes)),
    );

    for (const output of outputs) expect(output).toEqual(outputs[0]);
    expect(new Set(outputs).size).toBe(outputs.length);
    expect(new Set(outputs.map((output) => output.sheets)).size).toBe(
      outputs.length,
    );
  });

  it('returns the document-only convenience result', async () => {
    const bytes = await createIndependentXlsx();

    await expect(parseXlsx(bytes)).resolves.toEqual(
      (await parseXlsxWithDiagnostics(bytes)).document,
    );
  });

  it.each(['tolerant', 'strict'] as const)(
    'rejects invalid ZIP data with a structured error in %s mode',
    async (errorMode) => {
      const action = parseXlsx(new Uint8Array([1, 2, 3, 4]), { errorMode });
      await expect(action).rejects.toMatchObject({
        diagnostic: {
          code: 'invalid-package',
          message: 'Failed to open XLSX OPC package',
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
      const error = await action.catch((reason: unknown) => reason);
      expect(error).toBeInstanceOf(XlsxParseError);
      expect((error as XlsxParseError).cause).toBeInstanceOf(Error);
    },
  );

  it.each(['tolerant', 'strict'] as const)(
    'enforces input limits with structured metadata in %s mode',
    async (errorMode) => {
      const bytes = await createIndependentXlsx();
      await expect(
        parseXlsx(bytes, {
          errorMode,
          limits: { maxInputBytes: bytes.byteLength - 1 },
        }),
      ).rejects.toMatchObject({
        cause: {
          actual: bytes.byteLength,
          limit: bytes.byteLength - 1,
          limitName: 'maxInputBytes',
          name: 'XlsxResourceLimitError',
        },
        diagnostic: {
          actual: bytes.byteLength,
          code: 'resource-limit-exceeded',
          limit: bytes.byteLength - 1,
          limitName: 'maxInputBytes',
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
    },
  );

  it('maps archive preflight limits to the public structured error', async () => {
    const bytes = await createIndependentXlsx();

    await expect(
      parseXlsx(bytes, { limits: { maxEntries: 1 } }),
    ).rejects.toMatchObject({
      cause: {
        limit: 1,
        limitName: 'maxEntries',
        name: 'XlsxResourceLimitError',
      },
      diagnostic: {
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxEntries',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('maps streamed workbook-table limits to the public structured error', async () => {
    const bytes = await createIndependentXlsx({
      'xl/sharedStrings.xml': `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si/><si/></sst>`,
    });

    await expect(
      parseXlsx(bytes, { limits: { maxSharedStrings: 1 } }),
    ).rejects.toMatchObject({
      cause: {
        actual: 2,
        limit: 1,
        limitName: 'maxSharedStrings',
        name: 'XlsxResourceLimitError',
        part: 'xl/sharedStrings.xml',
      },
      diagnostic: {
        actual: 2,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxSharedStrings',
        part: 'xl/sharedStrings.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('rejects malformed shared strings through the public parser', async () => {
    const bytes = await createIndependentXlsx({
      'xl/sharedStrings.xml': `<sst xmlns="urn:wrong"><si><t>bad</t></si></sst>`,
    });

    await expect(parseXlsx(bytes)).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        message: 'Shared-string element has an unsupported namespace',
        part: 'xl/sharedStrings.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it.each([
    { errorMode: 'strict' },
    { errorMode: 'tolerant' },
    { displayTextMode: 'none' },
    { displayTextMode: 'supported' },
    { imageMode: 'base64' },
    { imageMode: 'blob' },
    { imageMode: 'both' },
    { imageMode: 'none' },
    { pivotCacheMode: 'metadata' },
    { pivotCacheMode: 'none' },
    { pivotCacheMode: 'records' },
  ] as const)('accepts valid runtime option %#', async (options) => {
    const bytes = await createIndependentXlsx();
    await expect(parseXlsx(bytes, options)).resolves.toMatchObject({
      sheets: [{ name: 'Sheet1' }],
    });
  });

  it.each([
    ['errorMode', 'recover', 'XLSX errorMode is invalid'],
    ['displayTextMode', 'all', 'XLSX displayTextMode is invalid'],
    ['imageMode', 'url', 'XLSX imageMode is invalid'],
    ['pivotCacheMode', 'all', 'XLSX pivotCacheMode is invalid'],
  ] as const)(
    'rejects invalid runtime option %s',
    async (name, value, message) => {
      const bytes = await createIndependentXlsx();
      await expect(parseXlsx(bytes, { [name]: value })).rejects.toThrow(
        message,
      );
    },
  );
});
