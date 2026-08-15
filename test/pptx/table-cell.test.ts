import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import { getTableCellParams } from '../../src/formats/pptx/internal/table';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(): PptxParserContext {
  return {
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    themeContent: xml({}),
  } as unknown as PptxParserContext;
}

describe('PPTX table cell structural attributes', () => {
  it('parses positive spans, active merge flags, and vertical alignment', () => {
    expect(
      getTableCellParams(
        xml({
          attrs: { gridSpan: '2', hMerge: '1', rowSpan: '3', vMerge: '1' },
          'a:tcPr': { attrs: { anchor: 'ctr' } },
        }),
        undefined,
        undefined,
        context(),
      ),
    ).toMatchObject({
      borders: {},
      colSpan: 2,
      hMerge: 1,
      rowSpan: 3,
      vAlign: 'mid',
      vMerge: 1,
    });
  });

  it.each([
    ['ctr', 'mid'],
    ['b', 'down'],
    ['t', 'up'],
    [undefined, 'up'],
    ['unsupported', 'up'],
  ] as const)('maps vertical anchor %j to %s', (anchor, expected) => {
    const cell = getTableCellParams(
      xml({ 'a:tcPr': { attrs: { ...(anchor ? { anchor } : {}) } } }),
      undefined,
      undefined,
      context(),
    );

    expect(cell.vAlign).toBe(expected);
  });

  it.each(['2x', '-1', '0', '1.5', 'Infinity', 'not-a-number'])(
    'omits invalid row span %j',
    (rowSpan) => {
      const cell = getTableCellParams(
        xml({ attrs: { rowSpan } }),
        undefined,
        undefined,
        context(),
      );

      expect(cell).not.toHaveProperty('rowSpan');
      const numbers = Object.values(cell).filter(
        (value): value is number => typeof value === 'number',
      );
      expect(numbers.every(Number.isFinite)).toBe(true);
    },
  );

  it.each(['2x', '-1', '0', '1.5', 'Infinity', 'not-a-number'])(
    'omits invalid column span %j',
    (gridSpan) => {
      const cell = getTableCellParams(
        xml({ attrs: { gridSpan } }),
        undefined,
        undefined,
        context(),
      );

      expect(cell).not.toHaveProperty('colSpan');
    },
  );

  it.each(['0', '2', '-1', '1.5', 'Infinity', 'not-a-number'])(
    'omits inactive or malformed merge flag %j',
    (merge) => {
      const cell = getTableCellParams(
        xml({ attrs: { hMerge: merge, vMerge: merge } }),
        undefined,
        undefined,
        context(),
      );

      expect(cell).not.toHaveProperty('hMerge');
      expect(cell).not.toHaveProperty('vMerge');
    },
  );
});
