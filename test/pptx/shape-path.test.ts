import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getShapePath } from '../../src/formats/pptx/internal/shape-path';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function guides(
  values: ReadonlyArray<readonly [name: string, formula: string]>,
): XmlLookupValue {
  return xml({
    'p:spPr': {
      'a:prstGeom': {
        'a:avLst': {
          'a:gd': values.map(([name, fmla]) => ({
            attrs: { fmla, name },
          })),
        },
      },
    },
  });
}

function expectFinitePath(path: string): void {
  expect(path).not.toBe('');
  expect(path).not.toMatch(/NaN|Infinity|undefined/);
  expect(path.trimStart()).toMatch(/^M/i);
}

describe('PowerPoint preset shape path safety', () => {
  it.each([
    'roundRect',
    'snipRoundRect',
    'triangle',
    'trapezoid',
    'star5',
    'pie',
    'chord',
    'frame',
    'blockArc',
  ])('ignores malformed adjustment formulas for %s', (shapeType) => {
    expectFinitePath(
      getShapePath(
        shapeType,
        120,
        80,
        guides([
          ['adj1', 'val 10000junk'],
          ['adj2', 'not-a-formula'],
        ]),
      ),
    );
  });

  it.each([
    ['not-a-number', Number.NaN, 80],
    ['positive infinity', Number.POSITIVE_INFINITY, 80],
    ['negative width', -120, 80],
    ['zero width', 0, 80],
    ['negative height', 120, -80],
    ['zero height', 120, 0],
    ['overflowing width', Number.MAX_VALUE, 80],
  ])('returns a finite degenerate path for %s', (_name, width, height) => {
    expectFinitePath(getShapePath('ellipse', width, height, xml({})));
  });

  it('keeps common default paths deterministic', () => {
    expect(getShapePath('rect', 120, 80, xml({}))).toBe(
      'M 0 0 L 120 0 L 120 80 L 0 80 Z',
    );
    expect(getShapePath('rtTriangle', 120, 80, xml({}))).toBe(
      'M 0 0 L 0 80 L 120 80 Z',
    );
    expect(getShapePath('bentConnector2', 120, 80, xml({}))).toBe(
      'M 120 0 L 120 80 L 0 80',
    );
  });
});
