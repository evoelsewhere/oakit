import { describe, expect, it, vi } from 'vitest';

import type {
  PptxSceneTableCell,
  PptxSceneTableElement,
} from '../../src/formats/pptx/scene-types';
import type { Table } from '../../src/formats/pptx/types';
import {
  verifyPowerPointTableElement,
  type PptxTableVerificationDependencies,
} from '../../src/formats/pptx/writer/table-verify';

function textCell(
  overrides: Partial<PptxSceneTableCell> = {},
): PptxSceneTableCell {
  return {
    borders: {
      bottom: { color: '#111111', style: 'solid', width: 1 },
      left: { color: '#222222', style: 'dotted', width: 2 },
      right: { color: '#333333', style: 'dashed', width: 3 },
      top: { color: '#444444', width: 4 },
    },
    colSpan: 2,
    fillColor: '#ABCDEF',
    rowSpan: 2,
    text: {
      body: {},
      paragraphs: [
        {
          children: [
            { key: 'run', text: 'Alpha', type: 'run' },
            { key: 'break', type: 'break' },
            { key: 'field', fieldType: 'x', text: 'Beta', type: 'field' },
          ],
          key: 'paragraph',
        },
      ],
    },
    ...overrides,
  };
}

function expectedTable(): PptxSceneTableElement {
  return {
    authored: { transform: { height: 80, width: 300, x: 10, y: 20 } },
    columns: [100, 200],
    key: 'table',
    resolved: { hidden: false },
    rows: [
      {
        cells: [
          textCell(),
          {
            borders: {},
            hMerge: true,
            text: textCell().text,
            vMerge: true,
          },
        ],
        height: 80,
      },
    ],
    type: 'table',
  };
}

function parsedBorder(
  borderColor: string,
  borderType: 'dashed' | 'dotted' | 'solid',
  borderWidth: number,
) {
  return { borderColor, borderType, borderWidth };
}

function generatedTable(): Table {
  return {
    borders: {},
    colWidths: [100, 200],
    data: [
      [
        {
          borders: {
            bottom: parsedBorder('#111111', 'solid', 1),
            left: parsedBorder('#222222', 'dotted', 2),
            right: parsedBorder('#333333', 'dashed', 3),
            top: parsedBorder('#444444', 'solid', 4),
          },
          colSpan: 2,
          fillColor: '#ABCDEF',
          rowSpan: 2,
          text: 'Alpha\nBeta',
          vAlign: 'up',
        },
        {
          borders: {},
          hMerge: 1,
          text: 'producer placeholder',
          vAlign: 'up',
          vMerge: 1,
        },
      ],
    ],
    height: 80,
    id: '2',
    left: 10,
    order: 0,
    rowHeights: [80],
    top: 20,
    type: 'table',
    width: 300,
  };
}

function dependencies() {
  const verifyTransform = vi.fn();
  const value: PptxTableVerificationDependencies = {
    expectedPointValue: (number) => number,
    plainText: (html) => html,
    textNodeValue: (node) => (node.type === 'break' ? '\n' : node.text),
    verifyTransform,
  };
  return { value, verifyTransform };
}

describe('native PowerPoint table verification', () => {
  it('verifies every table field and ignores continuation placeholder text', () => {
    const expected = expectedTable();
    const generated = generatedTable();
    const { value, verifyTransform } = dependencies();

    expect(() =>
      verifyPowerPointTableElement(generated, expected, 1, 2, value),
    ).not.toThrow();
    expect(verifyTransform).toHaveBeenCalledWith(
      generated,
      expected,
      'slide 2, element 3',
    );
  });

  it.each([undefined, { type: 'shape' }])(
    'rejects missing or non-table generated element %#',
    (generated) => {
      expect(() =>
        verifyPowerPointTableElement(
          generated as never,
          expectedTable(),
          1,
          2,
          dependencies().value,
        ),
      ).toThrow('Generated PowerPoint table missing at slide 2, element 3');
    },
  );

  it.each([
    ['columns', (value: Table) => (value.colWidths = [101, 199])],
    ['rows', (value: Table) => (value.rowHeights = [79])],
    ['row count', (value: Table) => value.data.push([])],
  ])('rejects a mismatched %s grid', (_name, mutate) => {
    const generated = generatedTable();
    mutate(generated);

    expect(() =>
      verifyPowerPointTableElement(
        generated,
        expectedTable(),
        0,
        0,
        dependencies().value,
      ),
    ).toThrow('Generated PowerPoint table grid mismatch at slide 1, element 1');
  });

  it('rejects missing rows and exact row cell counts', () => {
    const missing = generatedTable();
    missing.data = [];
    expect(() =>
      verifyPowerPointTableElement(
        missing,
        expectedTable(),
        0,
        0,
        dependencies().value,
      ),
    ).toThrow('Generated PowerPoint table grid mismatch');

    const short = generatedTable();
    short.data[0] = [short.data[0]?.[0] as Table['data'][number][number]];
    expect(() =>
      verifyPowerPointTableElement(
        short,
        expectedTable(),
        1,
        2,
        dependencies().value,
      ),
    ).toThrow(
      'Generated PowerPoint table row mismatch at slide 2, element 3, row 1',
    );
  });

  it.each([
    [
      'fill',
      (cell: Table['data'][number][number]) => (cell.fillColor = '#000000'),
    ],
    [
      'column span',
      (cell: Table['data'][number][number]) => (cell.colSpan = 3),
    ],
    ['row span', (cell: Table['data'][number][number]) => (cell.rowSpan = 3)],
    [
      'horizontal merge',
      (cell: Table['data'][number][number]) => (cell.hMerge = 1),
    ],
    [
      'vertical merge',
      (cell: Table['data'][number][number]) => (cell.vMerge = 1),
    ],
  ])('rejects a mismatched cell %s', (_name, mutate) => {
    const generated = generatedTable();
    mutate(generated.data[0]?.[0] as Table['data'][number][number]);

    expect(() =>
      verifyPowerPointTableElement(
        generated,
        expectedTable(),
        0,
        0,
        dependencies().value,
      ),
    ).toThrow(
      'Generated PowerPoint table cell mismatch at slide 1, element 1, row 1, cell 1',
    );
  });

  it('rejects a missing cell and mismatched origin text', () => {
    const missing = generatedTable();
    if (missing.data[0] === undefined) throw new Error('Expected row');
    missing.data[0][1] = undefined as never;
    expect(() =>
      verifyPowerPointTableElement(
        missing,
        expectedTable(),
        0,
        0,
        dependencies().value,
      ),
    ).toThrow('Generated PowerPoint table cell missing');

    const text = generatedTable();
    const origin = text.data[0]?.[0];
    if (origin === undefined) throw new Error('Expected origin');
    origin.text = 'wrong';
    expect(() =>
      verifyPowerPointTableElement(
        text,
        expectedTable(),
        0,
        0,
        dependencies().value,
      ),
    ).toThrow('Generated PowerPoint table text mismatch');
  });

  it.each(['bottom', 'left', 'right', 'top'] as const)(
    'rejects missing, unexpected, and mismatched %s borders',
    (direction) => {
      const missing = generatedTable();
      delete (missing.data[0]?.[0] as Table['data'][number][number]).borders[
        direction
      ];
      expect(() =>
        verifyPowerPointTableElement(
          missing,
          expectedTable(),
          0,
          0,
          dependencies().value,
        ),
      ).toThrow(`Generated PowerPoint table border mismatch`);

      for (const key of ['borderColor', 'borderWidth', 'borderType'] as const) {
        const mismatch = generatedTable();
        const border = mismatch.data[0]?.[0]?.borders[direction];
        if (border === undefined) throw new Error('Expected border');
        Object.assign(border, {
          [key]:
            key === 'borderColor'
              ? '#FFFFFF'
              : key === 'borderWidth'
                ? 9
                : border.borderType === 'solid'
                  ? 'dashed'
                  : 'solid',
        });
        expect(() =>
          verifyPowerPointTableElement(
            mismatch,
            expectedTable(),
            0,
            0,
            dependencies().value,
          ),
        ).toThrow(`Generated PowerPoint table border mismatch`);
      }

      const unexpected = generatedTable();
      const expected = expectedTable();
      const continuation = expected.rows[0]?.cells[1];
      if (continuation === undefined) throw new Error('Expected continuation');
      continuation.borders = {};
      const generatedContinuation = unexpected.data[0]?.[1];
      if (generatedContinuation === undefined) throw new Error('Expected cell');
      generatedContinuation.borders[direction] = parsedBorder(
        '#000000',
        'solid',
        1,
      );
      expect(() =>
        verifyPowerPointTableElement(
          unexpected,
          expected,
          0,
          0,
          dependencies().value,
        ),
      ).toThrow(`Generated PowerPoint table border mismatch`);
    },
  );
});
