import { describe, expect, it } from 'vitest';

import {
  transformXlsxStructuralCell,
  transformXlsxStructuralRange,
  type XlsxStructuralReferenceOperation,
} from '../../src/formats/xlsx/roundtrip/structural-reference';
import type { XlsxRange } from '../../src/formats/xlsx/types';

function range(reference: string, start: number, end: number): XlsxRange {
  return {
    end: { column: 1, row: end },
    reference,
    start: { column: 1, row: start },
  };
}

describe('XLSX structural reference transforms', () => {
  it.each([
    [1, 'A1'],
    [2, 'A4'],
    [3, 'A5'],
    [4, 'A6'],
  ] as const)('inserts rows at coordinate %s', (row, address) => {
    expect(
      transformXlsxStructuralCell(row, 1, {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ address, column: 1, row: Number(address.slice(1)) });
  });

  it.each([
    [1, 'A1'],
    [2, null],
    [3, null],
    [4, 'A2'],
    [5, 'A3'],
  ] as const)('deletes rows at coordinate %s', (row, address) => {
    expect(
      transformXlsxStructuralCell(row, 1, {
        count: 2,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toEqual(
      address === null
        ? null
        : { address, column: 1, row: Number(address.slice(1)) },
    );
  });

  it.each([
    [1, 1, 'A1'],
    [2, 2, null],
    [1, 2, 'A1'],
    [2, 3, null],
    [1, 4, 'A1:A2'],
    [4, 5, 'A2:A3'],
    [3, 4, 'A2'],
    [3, 5, 'A2:A3'],
    [5, 6, 'A3:A4'],
  ] as const)(
    'clips deleted row interval %s:%s',
    (start, end, expectedReference) => {
      expect(
        transformXlsxStructuralRange(range(`A${start}:A${end}`, start, end), {
          count: 2,
          index: 2,
          kind: 'delete-rows',
        })?.reference ?? null,
      ).toBe(expectedReference);
    },
  );

  it('expands spanning ranges and transforms columns symmetrically', () => {
    expect(
      transformXlsxStructuralCell(2, 3, {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ address: 'C4', column: 3, row: 4 });
    expect(
      transformXlsxStructuralRange(range('A1:A2', 1, 2), {
        count: 2,
        index: 4,
        kind: 'delete-rows',
      })?.reference,
    ).toBe('A1:A2');
    expect(
      transformXlsxStructuralRange(range('A1:A1', 1, 1), {
        count: 2,
        index: 3,
        kind: 'insert-rows',
      })?.reference,
    ).toBe('A1');
    expect(
      transformXlsxStructuralRange(range('A1:A4', 1, 4), {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({
      end: { column: 1, row: 6 },
      reference: 'A1:A6',
      start: { column: 1, row: 1 },
    });
    const source: XlsxRange = {
      end: { column: 5, row: 2 },
      reference: 'B2:E2',
      start: { column: 2, row: 2 },
    };
    const operations: XlsxStructuralReferenceOperation[] = [
      { count: 2, index: 3, kind: 'insert-columns' },
      { count: 2, index: 2, kind: 'delete-columns' },
    ];
    const inserted = transformXlsxStructuralRange(source, operations[0]!)!;
    expect(inserted.reference).toBe('B2:G2');
    expect(transformXlsxStructuralRange(inserted, operations[1]!)).toEqual({
      end: { column: 5, row: 2 },
      reference: 'B2:E2',
      start: { column: 2, row: 2 },
    });
    expect(
      transformXlsxStructuralCell(2, 3, {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ address: 'D2', column: 4, row: 2 });
    expect(
      transformXlsxStructuralCell(2, 2, {
        count: 1,
        index: 2,
        kind: 'delete-columns',
      }),
    ).toBeNull();
  });
});
