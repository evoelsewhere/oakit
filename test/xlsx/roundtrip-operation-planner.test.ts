import { describe, expect, it } from 'vitest';

import { resolveXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { canonicalXlsxJson } from '../../src/formats/xlsx/roundtrip/canonical-json';
import { canonicalXlsxSha256 } from '../../src/formats/xlsx/roundtrip/digest';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import {
  replayXlsxCellOperations,
  xlsxCellTargetState,
} from '../../src/formats/xlsx/roundtrip/operation-planner';
import { validateXlsxCellOperations } from '../../src/formats/xlsx/roundtrip/operation-validation';
import { readXlsxRoundTrip } from '../../src/formats/xlsx/roundtrip/read-snapshot';
import type {
  XlsxEditOperation,
  XlsxRoundTripDocument,
} from '../../src/formats/xlsx/roundtrip/types';
import {
  defaultXlsxWriteLimits,
  resolveXlsxWriteLimits,
} from '../../src/formats/xlsx/roundtrip/write-limits';
import type { XlsxWorksheet } from '../../src/formats/xlsx/types';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const writeLimits = defaultXlsxWriteLimits();
const readerLimits = resolveXlsxResourceLimits();
const ERROR_CODES = [
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
] as const;
const UNSUPPORTED_OPERATION_KINDS = [
  'add-worksheet',
  'delete-columns',
  'delete-rows',
  'delete-worksheet',
  'insert-columns',
  'insert-rows',
  'rename-worksheet',
  'set-column',
  'set-hyperlink',
  'set-row',
] as const;

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX operation validation to fail');
}

async function captureAsync(
  action: () => Promise<unknown>,
): Promise<XlsxWriteError> {
  try {
    await action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX operation planning to fail');
}

function worksheet(document: XlsxRoundTripDocument): XlsxWorksheet & {
  key: string;
} {
  const sheet = document.sheets[0]!;
  expect(sheet.kind).toBe('worksheet');
  return sheet as XlsxWorksheet & { key: string };
}

function cellOperation(
  document: XlsxRoundTripDocument,
  overrides: Partial<Extract<XlsxEditOperation, { kind: 'set-cell' }>> = {},
): Extract<XlsxEditOperation, { kind: 'set-cell' }> {
  return {
    cell: 'A1',
    content: { kind: 'value', value: { kind: 'text', text: 'updated' } },
    kind: 'set-cell',
    operationId: 'edit-1',
    sheetKey: worksheet(document).key,
    ...overrides,
  };
}

describe('XLSX cell operation validation', () => {
  it('normalizes every supported scalar payload and formula', () => {
    const common = {
      cell: 'A1',
      operationId: 'agent:edit_1',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    const operations = validateXlsxCellOperations(
      [
        {
          ...common,
          content: { kind: 'value', value: { kind: 'number', value: -0 } },
          kind: 'set-cell',
        },
        {
          ...common,
          content: { kind: 'value', value: { kind: 'boolean', value: false } },
          kind: 'set-cell',
          operationId: 'boolean',
        },
        {
          ...common,
          content: { kind: 'value', value: { code: '#N/A', kind: 'error' } },
          kind: 'set-cell',
          operationId: 'error',
        },
        {
          ...common,
          content: { kind: 'formula', expression: 'SUM(A2:A3)' },
          ifMatch: 'b'.repeat(64),
          kind: 'set-cell',
          operationId: 'formula',
        },
        { ...common, kind: 'clear-cell', operationId: 'clear' },
      ],
      writeLimits,
      readerLimits,
    );
    expect(operations).toHaveLength(5);
    expect(operations[0]).toMatchObject({
      content: { value: { kind: 'number', value: 0 } },
    });
    expect(
      Object.is(
        (operations[0] as { content: { value: { value: number } } }).content
          .value.value,
        -0,
      ),
    ).toBe(false);
    expect(operations[3]).toMatchObject({
      content: { expression: 'SUM(A2:A3)', kind: 'formula' },
      ifMatch: 'b'.repeat(64),
    });
  });

  it.each([
    [null, 'XLSX round-trip operations must be an array'],
    [[null], 'XLSX operation shape is invalid'],
    [[[]], 'XLSX operation shape is invalid'],
    [[Object.create(null)], 'XLSX operation shape is invalid'],
    [[{ kind: 'wat', operationId: 'one' }], 'XLSX operation kind is invalid'],
    [
      [
        {
          cell: 'a1',
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation cell reference is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          extra: true,
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX clear-cell operation shape is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          kind: 'clear-cell',
          operationId: '-bad',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation ID is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: 'Sheet1',
        },
      ],
      'XLSX operation sheet key is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          ifMatch: 'A'.repeat(64),
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation precondition hash is invalid',
    ],
  ] as const)('rejects invalid operation contract %#', (value, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(value, writeLimits, readerLimits),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    ['one!', 'XLSX operation ID is invalid'],
    [`x${'a'.repeat(64)}`, 'XLSX operation precondition hash is invalid'],
    [`${'a'.repeat(64)}x`, 'XLSX operation precondition hash is invalid'],
    [`xxlsx:sheet:${'a'.repeat(32)}`, 'XLSX operation sheet key is invalid'],
    [`xlsx:sheet:${'a'.repeat(32)}x`, 'XLSX operation sheet key is invalid'],
  ] as const)(
    'requires whole-string operation identities %#',
    (value, message) => {
      const operation: Record<string, unknown> = {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'one',
        sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
      };
      if (message.includes('ID')) operation.operationId = value;
      else if (message.includes('precondition')) operation.ifMatch = value;
      else operation.sheetKey = value;
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe(message);
    },
  );

  it.each(ERROR_CODES)('accepts typed error value %s', (code) => {
    const operation = validateXlsxCellOperations(
      [
        {
          cell: 'A1',
          content: { kind: 'value', value: { code, kind: 'error' } },
          kind: 'set-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      writeLimits,
      readerLimits,
    )[0];
    expect(operation).toMatchObject({
      content: { value: { code, kind: 'error' } },
    });
  });

  it.each(UNSUPPORTED_OPERATION_KINDS)(
    'reports recognized unsupported operation %s',
    (kind) => {
      const error = capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              kind,
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass: kind,
        message: `XLSX operation ${kind} is not supported by this profile`,
        operationId: 'one',
      });
    },
  );

  it('requires every operation field while allowing only ifMatch as optional', () => {
    const clear: Record<string, unknown> = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'one',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    for (const key of ['cell', 'sheetKey']) {
      const operation = { ...clear };
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX clear-cell operation shape is invalid');
    }
    const set = {
      ...clear,
      content: { kind: 'value', value: { kind: 'number', value: 1 } },
      kind: 'set-cell',
    };
    for (const key of ['cell', 'content', 'sheetKey']) {
      const operation = { ...set } as Record<string, unknown>;
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX set-cell operation shape is invalid');
    }
    const style = { ...clear, kind: 'set-cell-style', style: {} };
    for (const key of ['cell', 'sheetKey', 'style']) {
      const operation = { ...style } as Record<string, unknown>;
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX set-cell-style operation shape is invalid');
    }
    const withMatch = validateXlsxCellOperations(
      [{ ...clear, ifMatch: 'b'.repeat(64) }],
      writeLimits,
      readerLimits,
    )[0];
    const withoutMatch = validateXlsxCellOperations(
      [clear],
      writeLimits,
      readerLimits,
    )[0];
    expect(withMatch).toHaveProperty('ifMatch', 'b'.repeat(64));
    expect(withoutMatch).not.toHaveProperty('ifMatch');
    expect(
      validateXlsxCellOperations(
        [{ ...style, ifMatch: 'c'.repeat(64) }],
        writeLimits,
        readerLimits,
      )[0],
    ).toHaveProperty('ifMatch', 'c'.repeat(64));
  });

  it.each([
    [
      { kind: 'value', value: { kind: 'text', runs: [], text: '' } },
      'rich-text',
    ],
    [
      {
        kind: 'value',
        value: {
          kind: 'date',
          normalized: '2024-01-01',
          precision: 'date',
          source: { kind: 'iso', value: '2024-01-01' },
        },
      },
      'date-value',
    ],
  ] as const)(
    'blocks valid but unsupported set-cell payload %#',
    (content, featureClass) => {
      const error = capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content,
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass,
        operationId: 'one',
      });
      expect(error.diagnostic.message).toBe(
        featureClass === 'rich-text'
          ? 'XLSX cell editing does not yet support rich text runs'
          : 'XLSX cell editing does not yet support date values',
      );
    },
  );

  it('distinguishes malformed cell and style payloads', () => {
    const base = {
      cell: 'A1',
      kind: 'set-cell',
      operationId: 'one',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    for (const content of [
      null,
      {},
      { kind: 'formula', expression: '' },
      { kind: 'formula', expression: '=A1' },
      { kind: 'value' },
    ]) {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [{ ...base, content }],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic.code,
      ).toBe('invalid-roundtrip-json');
    }
    for (const style of [null, [], { extra: true }]) {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [
              {
                cell: 'A1',
                kind: 'set-cell-style',
                operationId: 'one',
                sheetKey: base.sheetKey,
                style,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'invalid-roundtrip-json',
        message: 'XLSX set-cell-style style shape is invalid',
        operationId: 'one',
      });
    }
    expect(
      validateXlsxCellOperations(
        [
          {
            cell: 'A1',
            kind: 'set-cell-style',
            operationId: 'one',
            sheetKey: base.sheetKey,
            style: {},
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'one',
        sheetKey: base.sheetKey,
        style: {},
      },
    ]);
  });

  it.each([
    [null, 'XLSX cell value shape is invalid'],
    [undefined, 'XLSX cell value shape is invalid'],
    [1, 'XLSX cell value shape is invalid'],
    [[], 'XLSX cell value shape is invalid'],
    [Object.create(null) as unknown, 'XLSX cell value shape is invalid'],
    [{ kind: 'text' }, 'XLSX text cell value shape is invalid'],
    [
      { extra: true, kind: 'text', text: 'x' },
      'XLSX text cell value shape is invalid',
    ],
    [{ kind: 'text', text: 1 }, 'XLSX text cell value is invalid'],
    [{ kind: 'number' }, 'XLSX number cell value is invalid'],
    [{ kind: 'number', value: '1' }, 'XLSX number cell value is invalid'],
    [{ kind: 'number', value: Infinity }, 'XLSX number cell value is invalid'],
    [{ kind: 'boolean' }, 'XLSX boolean cell value is invalid'],
    [{ kind: 'boolean', value: 0 }, 'XLSX boolean cell value is invalid'],
    [{ code: '#BAD!', kind: 'error' }, 'XLSX error cell value is invalid'],
    [{ kind: 'error' }, 'XLSX error cell value is invalid'],
    [{ kind: 'unknown' }, 'XLSX cell value kind is invalid'],
  ] as const)('rejects malformed cell value %#', (value, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content: { kind: 'value', value },
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    [undefined, 'XLSX set-cell content shape is invalid'],
    [1, 'XLSX set-cell content shape is invalid'],
    [[], 'XLSX set-cell content shape is invalid'],
    [Object.create(null) as unknown, 'XLSX set-cell content shape is invalid'],
    [{ kind: 'formula' }, 'XLSX set-cell formula is invalid'],
    [{ expression: 1, kind: 'formula' }, 'XLSX set-cell formula is invalid'],
    [
      { expression: 'A1', extra: true, kind: 'formula' },
      'XLSX set-cell formula is invalid',
    ],
    [
      { extra: true, kind: 'value', value: { kind: 'number', value: 1 } },
      'XLSX set-cell value content shape is invalid',
    ],
    [
      { kind: 'value', other: true },
      'XLSX set-cell value content shape is invalid',
    ],
    [{ kind: 'other' }, 'XLSX set-cell content kind is invalid'],
  ] as const)('rejects malformed set-cell content %#', (content, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content,
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('rejects duplicate operation IDs before replay', () => {
    const operation = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'same',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    const error = capture(() =>
      validateXlsxCellOperations(
        [operation, operation],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-roundtrip-json',
      operationId: 'same',
    });
    expect(error.diagnostic.message).toBe('XLSX operation IDs must be unique');
  });

  it('enforces operation, formula, and text budgets at exact boundaries', () => {
    const sheetKey = `xlsx:sheet:${'a'.repeat(32)}`;
    const clear = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'one',
      sheetKey,
    };
    const exactBytes = new TextEncoder().encode(
      JSON.stringify(clear, Object.keys(clear).sort()),
    ).byteLength;
    expect(() =>
      validateXlsxCellOperations(
        [clear],
        { ...writeLimits, maxOperationBytes: exactBytes, maxOperations: 1 },
        readerLimits,
      ),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [clear],
          { ...writeLimits, maxOperationBytes: exactBytes - 1 },
          readerLimits,
        ),
      ).diagnostic,
    ).toMatchObject({
      limitName: 'maxOperationBytes',
      message: 'XLSX operation exceeds its byte limit',
    });
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [clear, { ...clear, operationId: 'two' }],
          { ...writeLimits, maxOperations: 1 },
          readerLimits,
        ),
      ).diagnostic.limitName,
    ).toBe('maxOperations');
    const secondClear = { ...clear, operationId: 'two' };
    const secondBytes = new TextEncoder().encode(
      canonicalXlsxJson(secondClear),
    ).byteLength;
    expect(() =>
      validateXlsxCellOperations(
        [clear, secondClear],
        {
          ...writeLimits,
          maxTotalOperationBytes: exactBytes + secondBytes,
        },
        readerLimits,
      ),
    ).not.toThrow();
    const totalByteError = capture(() =>
      validateXlsxCellOperations(
        [clear, secondClear],
        {
          ...writeLimits,
          maxTotalOperationBytes: exactBytes + secondBytes - 1,
        },
        readerLimits,
      ),
    );
    expect(totalByteError.diagnostic).toMatchObject({
      actual: exactBytes + secondBytes,
      code: 'resource-limit-exceeded',
      limit: exactBytes + secondBytes - 1,
      limitName: 'maxTotalOperationBytes',
      message: 'XLSX operations exceed their total byte limit',
      operationId: 'two',
    });

    const formula = {
      ...clear,
      content: { kind: 'formula', expression: 'AB' },
      kind: 'set-cell',
    };
    expect(() =>
      validateXlsxCellOperations([formula], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([formula], writeLimits, {
          ...readerLimits,
          maxFormulaCharacters: 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      limitName: 'maxFormulaCharacters',
      message: 'XLSX operation formula exceeds its character limit',
    });
    const formulaTwo = { ...formula, operationId: 'two' };
    expect(() =>
      validateXlsxCellOperations([formula, formulaTwo], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
        maxTotalFormulaCharacters: 4,
      }),
    ).not.toThrow();
    const totalFormulaError = capture(() =>
      validateXlsxCellOperations([formula, formulaTwo], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
        maxTotalFormulaCharacters: 3,
      }),
    );
    expect(totalFormulaError.diagnostic).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTotalFormulaCharacters',
      message: 'XLSX operations exceed their total formula character limit',
      operationId: 'two',
    });

    const text = {
      ...clear,
      content: { kind: 'value', value: { kind: 'text', text: 'éé' } },
      kind: 'set-cell',
    };
    expect(() =>
      validateXlsxCellOperations([text], writeLimits, {
        ...readerLimits,
        maxTextCharacters: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([text], writeLimits, {
          ...readerLimits,
          maxTextCharacters: 1,
        }),
      ).diagnostic.limitName,
    ).toBe('maxTextCharacters');
    const singleCharacterText = {
      ...text,
      content: { kind: 'value', value: { kind: 'text', text: 'x' } },
    };
    const secondText = { ...singleCharacterText, operationId: 'two' };
    expect(() =>
      validateXlsxCellOperations(
        [singleCharacterText, secondText],
        writeLimits,
        { ...readerLimits, maxTextCharacters: 2 },
      ),
    ).not.toThrow();
    const totalTextError = capture(() =>
      validateXlsxCellOperations(
        [singleCharacterText, secondText],
        writeLimits,
        { ...readerLimits, maxTextCharacters: 1 },
      ),
    );
    expect(totalTextError.diagnostic).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxTextCharacters',
      message: 'XLSX operations exceed their text character limit',
      operationId: 'two',
    });
  });
});

describe('XLSX cell operation planner', () => {
  it('replays ordered set and clear operations without mutating inputs', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const before = JSON.stringify(snapshot.document);
    const operations: XlsxEditOperation[] = [
      cellOperation(snapshot.document, {
        content: { kind: 'value', value: { kind: 'number', value: 7 } },
      }),
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'edit-2',
        sheetKey: worksheet(snapshot.document).key,
      },
      cellOperation(snapshot.document, {
        content: { kind: 'formula', expression: '1+2' },
        operationId: 'edit-3',
      }),
    ];
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      operations,
      writeLimits,
      readerLimits,
    );
    expect(JSON.stringify(snapshot.document)).toBe(before);
    expect(plan.operations).not.toBe(operations);
    expect(plan.impacts).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell',
        operationId: 'edit-1',
        sheetKey: worksheet(snapshot.document).key,
      },
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'edit-2',
        sheetKey: worksheet(snapshot.document).key,
      },
      {
        cell: 'A1',
        kind: 'set-cell',
        operationId: 'edit-3',
        sheetKey: worksheet(snapshot.document).key,
      },
    ]);
    expect(worksheet(plan.document).rows[0]?.cells[0]?.content).toEqual({
      cached: { kind: 'missing' },
      formula: { expression: '1+2', kind: 'normal' },
      kind: 'formula',
    });
    expect(plan.stateHash).toBe(await canonicalXlsxSha256(plan.document));
  });

  it('applies existing normalized styles and rejects append requests', async () => {
    const styles = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({ 'xl/styles.xml': styles }),
    );
    const sheet = worksheet(snapshot.document);
    const targetStyle = snapshot.document.styles[1]!;
    const operation = {
      cell: 'A1',
      kind: 'set-cell-style' as const,
      operationId: 'style-1',
      sheetKey: sheet.key,
      style: targetStyle,
    };
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [operation],
      writeLimits,
      readerLimits,
    );
    expect(sheet.rows[0]!.cells[0]!.style).toBeUndefined();
    expect(worksheet(plan.document).rows[0]!.cells[0]!.style).toBe(1);
    expect(plan.impacts).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-1',
        sheetKey: sheet.key,
      },
    ]);
    const defaultPlan = await replayXlsxCellOperations(
      snapshot.document,
      [
        {
          ...operation,
          operationId: 'style-0',
          style: snapshot.document.styles[0]!,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(defaultPlan.document).rows[0]!.cells[0]!.style).toBe(0);
    const error = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [{ ...operation, style: { font: { italic: true } } }],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      cell: 'A1',
      code: 'unsupported-edit-operation',
      featureClass: 'append-style',
      message:
        'XLSX set-cell-style currently requires an existing normalized style',
      operationId: 'style-1',
      sheetKey: sheet.key,
    });
  });

  it('checks ifMatch against the sequential target state', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const sourceSheet = worksheet(snapshot.document);
    const sourceCell = sourceSheet.rows[0]!.cells[0]!;
    const sourceMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(sourceSheet, sourceCell),
    );
    const first = cellOperation(snapshot.document, { ifMatch: sourceMatch });
    const firstPlan = await replayXlsxCellOperations(
      snapshot.document,
      [first],
      writeLimits,
      readerLimits,
    );
    const updatedSheet = worksheet(firstPlan.document);
    const updatedMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(updatedSheet, updatedSheet.rows[0]!.cells[0]!),
    );
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [
        first,
        cellOperation(snapshot.document, {
          content: { kind: 'value', value: { kind: 'boolean', value: true } },
          ifMatch: updatedMatch,
          operationId: 'edit-2',
        }),
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(plan.document).rows[0]?.cells[0]?.content).toEqual({
      kind: 'value',
      value: { kind: 'boolean', value: true },
    });

    const error = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [first, { ...first, operationId: 'edit-2' }],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      cell: 'A1',
      code: 'operation-precondition-failed',
      operationId: 'edit-2',
      sheetKey: sourceSheet.key,
    });
    expect(error.diagnostic.message).toBe(
      'XLSX operation precondition does not match the target cell',
    );
  });

  it('blocks missing targets and resolves row and column coordinates exactly', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
          <sheetData>
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row>
          </sheetData>
        </worksheet>`,
      }),
    );
    const cases = [
      cellOperation(snapshot.document, {
        sheetKey: `xlsx:sheet:${'f'.repeat(32)}`,
      }),
      cellOperation(snapshot.document, { cell: 'A2' }),
      cellOperation(snapshot.document, { cell: 'C1' }),
    ];
    const features = ['worksheet', 'missing-cell', 'missing-cell'];
    const messages = [
      'XLSX operation sheet key does not exist in the snapshot',
      'XLSX cell operation requires an existing explicit source cell',
      'XLSX cell operation requires an existing explicit source cell',
    ];
    for (const [index, operation] of cases.entries()) {
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [operation],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'preservation-conflict',
        featureClass: features[index],
        operationId: 'edit-1',
      });
      expect(error.diagnostic.message).toBe(messages[index]);
    }
  });

  it('applies merged-range geometry only inside the range and only at its anchor', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
          <sheetData>
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>1</v></c><c r="D1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c><c r="C2"><v>1</v></c><c r="D2"><v>1</v></c></row>
            <row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>1</v></c><c r="C3"><v>1</v></c><c r="D3"><v>1</v></c></row>
            <row r="4"><c r="A4"><v>1</v></c><c r="B4"><v>1</v></c><c r="D4"><v>1</v></c></row>
          </sheetData>
          <mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>
        </worksheet>`,
      }),
    );
    for (const cell of ['A2', 'D2', 'B1', 'B4', 'B2']) {
      await expect(
        replayXlsxCellOperations(
          snapshot.document,
          [cellOperation(snapshot.document, { cell })],
          writeLimits,
          readerLimits,
        ),
      ).resolves.toMatchObject({
        impacts: [{ cell, operationId: 'edit-1' }],
      });
    }
    for (const cell of ['C2', 'B3']) {
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [cellOperation(snapshot.document, { cell })],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        cell,
        code: 'preservation-conflict',
        featureClass: 'merged-cell',
        message: 'XLSX cell operation cannot target a non-anchor merged cell',
      });
    }
  });

  it('rejects chart-sheet targets and duplicate snapshot keys', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const chartDocument: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [
        {
          index: 0,
          key: worksheet(snapshot.document).key,
          kind: 'chart-sheet',
          name: 'Chart',
          payload: 'full-sheet',
          state: 'visible',
        },
      ],
    };
    const chartError = await captureAsync(() =>
      replayXlsxCellOperations(
        chartDocument,
        [cellOperation(snapshot.document)],
        writeLimits,
        readerLimits,
      ),
    );
    expect(chartError.diagnostic.featureClass).toBe('chart-sheet');
    expect(chartError.diagnostic.message).toBe(
      'XLSX cell operation cannot target a chart sheet',
    );

    const duplicate: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [
        snapshot.document.sheets[0]!,
        { ...snapshot.document.sheets[0]!, index: 1 },
      ],
    };
    const duplicateError = await captureAsync(() =>
      replayXlsxCellOperations(duplicate, [], writeLimits, readerLimits),
    );
    expect(duplicateError.diagnostic).toMatchObject({
      code: 'snapshot-integrity-failed',
      message: 'XLSX snapshot sheet keys must be unique',
      objectKey: worksheet(snapshot.document).key,
    });
  });

  it('returns an isolated literal-equal state for an empty operation list', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [],
      resolveXlsxWriteLimits(undefined),
      readerLimits,
    );
    expect(plan.document).toEqual(snapshot.document);
    expect(plan.document).not.toBe(snapshot.document);
    expect(plan.impacts).toEqual([]);
    expect(plan.stateHash).toBe(snapshot.stateHash);
  });
});
