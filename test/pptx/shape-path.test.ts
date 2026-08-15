import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getShapePath } from '../../src/formats/pptx/internal/shape-path';

const PRESET_SHAPE_TYPES = [
  'accentBorderCallout1',
  'accentBorderCallout2',
  'accentBorderCallout3',
  'accentCallout1',
  'accentCallout2',
  'accentCallout3',
  'actionButtonBackPrevious',
  'actionButtonBeginning',
  'actionButtonBlank',
  'actionButtonDocument',
  'actionButtonEnd',
  'actionButtonForwardNext',
  'actionButtonHelp',
  'actionButtonHome',
  'actionButtonInformation',
  'actionButtonMovie',
  'actionButtonReturn',
  'actionButtonSound',
  'arc',
  'bentArrow',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'bentUpArrow',
  'bevel',
  'blockArc',
  'borderCallout1',
  'borderCallout2',
  'borderCallout3',
  'bracePair',
  'bracketPair',
  'callout1',
  'callout2',
  'callout3',
  'can',
  'chartPlus',
  'chartStar',
  'chartX',
  'chevron',
  'chord',
  'circularArrow',
  'cloud',
  'cloudCallout',
  'corner',
  'cornerTabs',
  'cube',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
  'curvedDownArrow',
  'curvedLeftArrow',
  'curvedRightArrow',
  'curvedUpArrow',
  'decagon',
  'diagStripe',
  'diamond',
  'dodecagon',
  'donut',
  'doubleWave',
  'downArrow',
  'downArrowCallout',
  'ellipse',
  'ellipseRibbon',
  'ellipseRibbon2',
  'flowChartAlternateProcess',
  'flowChartCollate',
  'flowChartConnector',
  'flowChartDecision',
  'flowChartDelay',
  'flowChartDisplay',
  'flowChartDocument',
  'flowChartExtract',
  'flowChartInputOutput',
  'flowChartInternalStorage',
  'flowChartMagneticDisk',
  'flowChartMagneticDrum',
  'flowChartMagneticTape',
  'flowChartManualInput',
  'flowChartManualOperation',
  'flowChartMerge',
  'flowChartMultidocument',
  'flowChartOfflineStorage',
  'flowChartOffpageConnector',
  'flowChartOnlineStorage',
  'flowChartOr',
  'flowChartPredefinedProcess',
  'flowChartPreparation',
  'flowChartPunchedCard',
  'flowChartPunchedTape',
  'flowChartSort',
  'flowChartSummingJunction',
  'flowChartTerminator',
  'foldedCorner',
  'folderCorner',
  'frame',
  'funnel',
  'gear6',
  'gear9',
  'halfFrame',
  'heart',
  'heptagon',
  'hexagon',
  'homePlate',
  'horizontalScroll',
  'irregularSeal1',
  'irregularSeal2',
  'leftArrow',
  'leftArrowCallout',
  'leftBrace',
  'leftBracket',
  'leftCircularArrow',
  'leftRightArrow',
  'leftRightArrowCallout',
  'leftRightCircularArrow',
  'leftRightRibbon',
  'leftRightUpArrow',
  'leftUpArrow',
  'lightningBolt',
  'line',
  'lineInv',
  'mathDivide',
  'mathEqual',
  'mathMinus',
  'mathMultiply',
  'mathNotEqual',
  'mathPlus',
  'moon',
  'noSmoking',
  'nonIsoscelesTrapezoid',
  'notchedRightArrow',
  'octagon',
  'parallelogram',
  'pentagon',
  'pie',
  'pieWedge',
  'plaque',
  'plaqueTabs',
  'plus',
  'quadArrow',
  'quadArrowCallout',
  'rect',
  'ribbon',
  'ribbon2',
  'rightArrow',
  'rightArrowCallout',
  'rightBrace',
  'rightBracket',
  'round1Rect',
  'round2DiagRect',
  'round2SameRect',
  'roundRect',
  'rtTriangle',
  'smileyFace',
  'snip1Rect',
  'snip2DiagRect',
  'snip2SameRect',
  'snipRoundRect',
  'squareTabs',
  'star10',
  'star12',
  'star16',
  'star24',
  'star32',
  'star4',
  'star5',
  'star6',
  'star7',
  'star8',
  'straightConnector1',
  'stripedRightArrow',
  'sun',
  'swooshArrow',
  'teardrop',
  'trapezoid',
  'triangle',
  'upArrow',
  'upArrowCallout',
  'upDownArrow',
  'upDownArrowCallout',
  'uturnArrow',
  'verticalScroll',
  'wave',
  'wedgeEllipseCallout',
  'wedgeRectCallout',
  'wedgeRoundRectCallout',
] as const;

const AUTHORED_GUIDES = guides([
  ['adj', 'val 25000'],
  ['adj1', 'val 10000'],
  ['adj2', 'val 20000'],
  ['adj3', 'val 30000'],
  ['adj4', 'val 40000'],
  ['adj5', 'val 12500'],
  ['adj6', 'val 15000'],
  ['adj7', 'val 17500'],
  ['adj8', 'val 20000'],
  ['hf', 'val 100000'],
  ['vf', 'val 100000'],
]);

const SINGLE_GUIDE_VARIANTS = [
  ['adj', '25000'],
  ['adj1', '10000'],
  ['adj2', '20000'],
  ['adj3', '30000'],
  ['adj4', '40000'],
  ['adj5', '12500'],
  ['adj6', '15000'],
  ['adj7', '17500'],
  ['adj8', '20000'],
  ['hf', '100000'],
  ['vf', '100000'],
] as const;

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

function singleGuide(name: string, value: string): XmlLookupValue {
  return xml({
    'p:spPr': {
      'a:prstGeom': {
        'a:avLst': {
          'a:gd': { attrs: { fmla: `val ${value}`, name } },
        },
      },
    },
  });
}

function rawGuides(values: object | object[]): XmlLookupValue {
  return xml({
    'p:spPr': {
      'a:prstGeom': {
        'a:avLst': {
          'a:gd': values,
        },
      },
    },
  });
}

function hashPath(path: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < path.length; index += 1) {
    hash = Math.imul(hash ^ path.charCodeAt(index), 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function expectFinitePath(path: string): void {
  expect(path).not.toBe('');
  expect(path).not.toMatch(/NaN|Infinity|undefined/);
  expect(path.trimStart()).toMatch(/^M/i);
}

describe('PowerPoint preset shape path safety', () => {
  it.each(PRESET_SHAPE_TYPES)(
    'renders finite default geometry for %s',
    (shapeType) => {
      expectFinitePath(getShapePath(shapeType, 120, 80, xml({})));
    },
  );

  it.each(PRESET_SHAPE_TYPES)(
    'keeps the locked geometry contract for %s',
    (shapeType) => {
      const paths: Record<string, string> = {
        authored: getShapePath(shapeType, 120, 80, AUTHORED_GUIDES),
        default: getShapePath(shapeType, 120, 80, xml({})),
      };
      for (const [name, value] of SINGLE_GUIDE_VARIANTS) {
        paths[name] = getShapePath(
          shapeType,
          120,
          80,
          singleGuide(name, value),
        );
      }
      for (const path of Object.values(paths)) expectFinitePath(path);

      expect(
        Object.fromEntries(
          Object.entries(paths).map(([name, path]) => [name, hashPath(path)]),
        ),
      ).toMatchSnapshot();
    },
  );

  it.each(PRESET_SHAPE_TYPES)(
    'renders finite authored geometry for %s',
    (shapeType) => {
      expectFinitePath(getShapePath(shapeType, 120, 80, AUTHORED_GUIDES));
    },
  );

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
    ['missing name', { attrs: { fmla: 'val 10000' } }],
    ['missing formula', { attrs: { name: 'adj1' } }],
    ['formula prefix', { attrs: { fmla: 'prefix val 10000', name: 'adj1' } }],
    ['formula suffix', { attrs: { fmla: 'val 10000 suffix', name: 'adj1' } }],
    [
      'unexpected sign character',
      { attrs: { fmla: 'val x10000', name: 'adj1' } },
    ],
    ['decimal', { attrs: { fmla: 'val 100.5', name: 'adj1' } }],
    ['exponent', { attrs: { fmla: 'val 1e4', name: 'adj1' } }],
    ['hexadecimal', { attrs: { fmla: 'val 0x10', name: 'adj1' } }],
    ['wrong operator', { attrs: { fmla: 'foo 10000', name: 'adj1' } }],
    ['empty integer', { attrs: { fmla: 'val ', name: 'adj1' } }],
    ['digit below zero', { attrs: { fmla: 'val /10000', name: 'adj1' } }],
    ['digit above nine', { attrs: { fmla: 'val :10000', name: 'adj1' } }],
    [
      'unsafe integer',
      { attrs: { fmla: 'val 9007199254740992', name: 'adj1' } },
    ],
  ])('falls back to default geometry for %s', (_name, guide) => {
    expect(getShapePath('triangle', 120, 80, rawGuides(guide))).toBe(
      getShapePath('triangle', 120, 80, xml({})),
    );
  });

  it.each([
    ['explicit plus', 'val +10000', 'val 10000'],
    ['negative value', 'val -10000', 'val -10000'],
    ['leading zeros', 'val 0010000', 'val 10000'],
  ])('accepts %s in an integer guide', (_name, formula, equivalent) => {
    const path = getShapePath(
      'triangle',
      120,
      80,
      rawGuides({ attrs: { fmla: formula, name: 'adj1' } }),
    );

    expect(path).toBe(
      getShapePath(
        'triangle',
        120,
        80,
        rawGuides({ attrs: { fmla: equivalent, name: 'adj1' } }),
      ),
    );
    expect(path).not.toBe(getShapePath('triangle', 120, 80, xml({})));
  });

  it('rejects a malformed guide when valid guides precede it', () => {
    const node = rawGuides([
      { attrs: { fmla: 'val 10000', name: 'adj1' } },
      { attrs: { fmla: 'val 20000 suffix', name: 'adj2' } },
    ]);

    expect(getShapePath('round2DiagRect', 120, 80, node)).toBe(
      getShapePath('round2DiagRect', 120, 80, xml({})),
    );
  });

  it.each([
    ['not-a-number', Number.NaN, 80, 'M 0 0 L 0 0 L 0 80 L 0 80 Z'],
    [
      'positive infinity',
      Number.POSITIVE_INFINITY,
      80,
      'M 0 0 L 0 0 L 0 80 L 0 80 Z',
    ],
    ['negative width', -120, 80, 'M 0 0 L 0 0 L 0 80 L 0 80 Z'],
    ['zero width', 0, 80, 'M 0 0 L 0 0 L 0 80 L 0 80 Z'],
    ['negative height', 120, -80, 'M 0 0 L 120 0 L 120 0 L 0 0 Z'],
    ['zero height', 120, 0, 'M 0 0 L 120 0 L 120 0 L 0 0 Z'],
    ['overflowing width', Number.MAX_VALUE, 80, 'M 0 0 L 0 0 L 0 80 L 0 80 Z'],
  ])(
    'returns the exact degenerate path for %s',
    (_name, width, height, expected) => {
      expect(getShapePath('ellipse', width, height, xml({}))).toBe(expected);
    },
  );

  it.each([
    ['smallest positive width', Number.MIN_VALUE],
    ['largest safe width', Number.MAX_SAFE_INTEGER],
  ])('preserves the %s dimension boundary', (_name, width) => {
    expect(getShapePath('rect', width, 80, xml({}))).toBe(
      `M 0 0 L ${width} 0 L ${width} 80 L 0 80 Z`,
    );
  });

  it('distinguishes zero, half, and full pie sweeps', () => {
    const zero = getShapePath(
      'pie',
      120,
      80,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 0'],
      ]),
    );
    const half = getShapePath(
      'pie',
      120,
      80,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 10800000'],
      ]),
    );
    const full = getShapePath(
      'pie',
      120,
      80,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 21600000'],
      ]),
    );

    expect(zero).toBe('M60,40 L120,40 A60,40 0 0,1 120,40 z');
    expect(half).toContain('A60,40 0 0,1 ');
    expect(full).toContain('A60,40 0 1,1 ');
  });

  it('keeps an open half arc on the short-arc boundary', () => {
    expect(
      getShapePath(
        'arc',
        120,
        80,
        guides([
          ['adj1', 'val 0'],
          ['adj2', 'val 10800000'],
        ]),
      ),
    ).toContain('A60,40 0 0,1 ');
  });

  it('renders different bounded tooth counts for gear presets', () => {
    const gear6 = getShapePath('gear6', 120, 80, xml({}));
    const gear9 = getShapePath('gear9', 120, 80, xml({}));

    expect(gear6).not.toBe(gear9);
    expect(gear6.match(/\bL\b/g)).toHaveLength(23);
    expect(gear9.match(/\bL\b/g)).toHaveLength(35);
    expectFinitePath(gear6);
    expectFinitePath(gear9);
  });

  it('preserves flowchart orientation without zero-value mirroring', () => {
    expect(getShapePath('flowChartMerge', 120, 80, xml({}))).toBe(
      'M 60 80 L 120 0 L 0 0 Z',
    );
    expect(getShapePath('flowChartManualInput', 120, 80, xml({}))).toBe(
      'M 0 16 L 0 80 L 120 80 L 120 0 Z',
    );
    expect(getShapePath('flowChartManualOperation', 120, 80, xml({}))).toBe(
      'M 96 80 L 120 0 L 0 0 L 24 80 Z',
    );
  });

  it.each([
    ['wide', 120, 80, 'M 20 0 L 0 80 L 100 80 L 120 0 Z'],
    [
      'tall',
      80,
      120,
      'M 13.333333333333332 0 L 0 120 L 66.66666666666667 120 L 80 0 Z',
    ],
    ['square', 100, 100, 'M 25 0 L 0 100 L 75 100 L 100 0 Z'],
  ])(
    'scales an authored parallelogram in a %s box',
    (_name, width, height, expected) => {
      expect(
        getShapePath(
          'parallelogram',
          width,
          height,
          singleGuide('adj', '25000'),
        ),
      ).toBe(expected);
    },
  );

  it.each([
    'hexagon',
    'star4',
    'star5',
    'star6',
    'star7',
    'star8',
    'star10',
    'star12',
    'star16',
    'star24',
    'star32',
    'frame',
    'donut',
    'noSmoking',
  ])('clamps authored adjustment bounds for %s', (shapeType) => {
    const lowerBound = getShapePath(
      shapeType,
      100,
      100,
      singleGuide('adj', '0'),
    );
    const upperBound = getShapePath(
      shapeType,
      100,
      100,
      singleGuide('adj', '50000'),
    );
    expect(
      getShapePath(shapeType, 100, 100, singleGuide('adj', '-10000')),
    ).toBe(lowerBound);
    expect(
      getShapePath(shapeType, 100, 100, singleGuide('adj', '100000')),
    ).toBe(upperBound);
    expectFinitePath(lowerBound);
    expectFinitePath(upperBound);
  });

  it.each(['star10', 'star12', 'star16', 'star24', 'star32'])(
    'applies equivalent single and array guides for %s',
    (shapeType) => {
      const adjusted = getShapePath(
        shapeType,
        120,
        80,
        singleGuide('adj', '25000'),
      );
      expect(adjusted).toBe(
        getShapePath(shapeType, 120, 80, guides([['adj', 'val 25000']])),
      );
      expect(adjusted).not.toBe(getShapePath(shapeType, 120, 80, xml({})));
    },
  );

  it('clamps half-frame adjustments to its aspect-ratio limits', () => {
    expect(
      getShapePath(
        'halfFrame',
        120,
        80,
        guides([
          ['adj1', 'val 0'],
          ['adj2', 'val 200000'],
        ]),
      ),
    ).toBe(
      getShapePath(
        'halfFrame',
        120,
        80,
        guides([
          ['adj1', 'val 0'],
          ['adj2', 'val 150000'],
        ]),
      ),
    );
    expect(
      getShapePath(
        'halfFrame',
        100,
        100,
        guides([
          ['adj1', 'val 100000'],
          ['adj2', 'val 25000'],
        ]),
      ),
    ).toBe('M 0,0 L 100,0 L 25,75 L 25,75 L 25,75 L 0,100 z');
  });

  it('clamps block-arc angles and thickness to their authored bounds', () => {
    const lowerBound = getShapePath(
      'blockArc',
      100,
      100,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 0'],
        ['adj3', 'val 0'],
      ]),
    );
    const upperBound = getShapePath(
      'blockArc',
      100,
      100,
      guides([
        ['adj1', 'val 21600000'],
        ['adj2', 'val 21600000'],
        ['adj3', 'val 50000'],
      ]),
    );
    expect(
      getShapePath(
        'blockArc',
        100,
        100,
        guides([
          ['adj1', 'val -60000'],
          ['adj2', 'val -60000'],
          ['adj3', 'val -10000'],
        ]),
      ),
    ).toBe(lowerBound);
    expect(
      getShapePath(
        'blockArc',
        100,
        100,
        guides([
          ['adj1', 'val 21660000'],
          ['adj2', 'val 21660000'],
          ['adj3', 'val 100000'],
        ]),
      ),
    ).toBe(upperBound);
    expectFinitePath(lowerBound);
    expectFinitePath(upperBound);
  });

  it('renders equal block-arc angles as a full ring sweep', () => {
    const path = getShapePath(
      'blockArc',
      120,
      80,
      guides([
        ['adj1', 'val 5400000'],
        ['adj2', 'val 5400000'],
        ['adj3', 'val 25000'],
      ]),
    );
    expect(path.match(/ L/g)?.length).toBeGreaterThan(700);
  });

  it.each([
    ['bracePair', 25000],
    ['bracketPair', 50000],
    ['diagStripe', 100000],
    ['teardrop', 200000],
    ['plaque', 50000],
    ['cube', 100000],
    ['bevel', 50000],
    ['foldedCorner', 50000],
    ['verticalScroll', 25000],
    ['horizontalScroll', 25000],
  ] as const)(
    'clamps the fixed authored adjustment range for %s',
    (shapeType, maximum) => {
      const lowerBound = getShapePath(
        shapeType,
        100,
        100,
        singleGuide('adj', '0'),
      );
      const upperBound = getShapePath(
        shapeType,
        100,
        100,
        singleGuide('adj', String(maximum)),
      );
      expect(
        getShapePath(shapeType, 100, 100, singleGuide('adj', '-10000')),
      ).toBe(lowerBound);
      expect(
        getShapePath(
          shapeType,
          100,
          100,
          singleGuide('adj', String(maximum * 2)),
        ),
      ).toBe(upperBound);
      expectFinitePath(lowerBound);
      expectFinitePath(upperBound);
    },
  );

  it('clamps the authored sun adjustment to its asymmetric range', () => {
    const lowerBound = getShapePath(
      'sun',
      100,
      100,
      singleGuide('adj', '12500'),
    );
    const upperBound = getShapePath(
      'sun',
      100,
      100,
      singleGuide('adj', '46875'),
    );
    expect(getShapePath('sun', 100, 100, singleGuide('adj', '0'))).toBe(
      lowerBound,
    );
    expect(getShapePath('sun', 100, 100, singleGuide('adj', '100000'))).toBe(
      upperBound,
    );
    expectFinitePath(lowerBound);
    expectFinitePath(upperBound);
  });

  it.each([
    ['leftBrace', '739518f4'],
    ['rightBrace', 'fe8631e9'],
  ] as const)(
    'clamps coupled authored adjustments for %s',
    (shapeType, coupledUpperBoundHash) => {
      const lowerBound = getShapePath(
        shapeType,
        100,
        100,
        guides([
          ['adj1', 'val 0'],
          ['adj2', 'val 0'],
        ]),
      );
      const upperSecondAdjustment = getShapePath(
        shapeType,
        100,
        100,
        guides([
          ['adj1', 'val 0'],
          ['adj2', 'val 100000'],
        ]),
      );
      expect(
        getShapePath(
          shapeType,
          100,
          100,
          guides([
            ['adj1', 'val -10000'],
            ['adj2', 'val -10000'],
          ]),
        ),
      ).toBe(lowerBound);
      expect(
        getShapePath(
          shapeType,
          100,
          100,
          guides([
            ['adj1', 'val 0'],
            ['adj2', 'val 200000'],
          ]),
        ),
      ).toBe(upperSecondAdjustment);
      const coupledUpperBound = getShapePath(
        shapeType,
        100,
        100,
        guides([
          ['adj1', 'val 100000'],
          ['adj2', 'val 75000'],
        ]),
      );
      expect(hashPath(coupledUpperBound)).toBe(coupledUpperBoundHash);
      expectFinitePath(lowerBound);
      expectFinitePath(upperSecondAdjustment);
    },
  );

  it.each([
    ['leftBracket', 100000],
    ['rightBracket', 150000],
  ] as const)(
    'clamps a tall authored %s to its geometry limit',
    (shapeType, effectiveMaximum) => {
      const lowerBound = getShapePath(
        shapeType,
        40,
        120,
        singleGuide('adj', '0'),
      );
      const upperBound = getShapePath(
        shapeType,
        40,
        120,
        singleGuide('adj', String(effectiveMaximum)),
      );
      expect(
        getShapePath(shapeType, 40, 120, singleGuide('adj', '-10000')),
      ).toBe(lowerBound);
      expect(
        getShapePath(shapeType, 40, 120, singleGuide('adj', '200000')),
      ).toBe(upperBound);
      expectFinitePath(lowerBound);
      expectFinitePath(upperBound);
    },
  );

  it('clamps corner adjustments independently for a wide box', () => {
    const lowerBound = getShapePath(
      'corner',
      120,
      80,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 0'],
      ]),
    );
    const upperBound = getShapePath(
      'corner',
      120,
      80,
      guides([
        ['adj1', 'val 100000'],
        ['adj2', 'val 150000'],
      ]),
    );
    expect(
      getShapePath(
        'corner',
        120,
        80,
        guides([
          ['adj1', 'val -10000'],
          ['adj2', 'val -10000'],
        ]),
      ),
    ).toBe(lowerBound);
    expect(
      getShapePath(
        'corner',
        120,
        80,
        guides([
          ['adj1', 'val 200000'],
          ['adj2', 'val 200000'],
        ]),
      ),
    ).toBe(upperBound);
    expectFinitePath(lowerBound);
    expectFinitePath(upperBound);
  });

  it('keeps a centered cloud callout finite', () => {
    const path = getShapePath(
      'cloudCallout',
      120,
      80,
      guides([
        ['adj1', 'val 0'],
        ['adj2', 'val 0'],
      ]),
    );

    expectFinitePath(path);
    expect(hashPath(path)).toBe('9d3cefdb');
  });

  it.each([
    [-50000, '3154bbac'],
    [50000, '64d18688'],
  ] as const)(
    'places a vertical cloud callout at adjustment %i',
    (adjustment, expectedHash) => {
      const path = getShapePath(
        'cloudCallout',
        120,
        80,
        guides([
          ['adj1', 'val 0'],
          ['adj2', `val ${adjustment}`],
        ]),
      );

      expectFinitePath(path);
      expect(hashPath(path)).toBe(expectedHash);
    },
  );

  it.each([
    [-4653, 'aa437488'],
    [0, '641e263f'],
    [1000, '96b0b84b'],
    [4653, '8507aae9'],
  ] as const)(
    'renders the smiley expression at adjustment %i',
    (adjustment, expectedHash) => {
      const path = getShapePath(
        'smileyFace',
        120,
        80,
        singleGuide('adj', String(adjustment)),
      );

      expectFinitePath(path);
      expect(hashPath(path)).toBe(expectedHash);
    },
  );

  it('clamps the smiley expression to its authored bounds', () => {
    const lowerBound = getShapePath(
      'smileyFace',
      120,
      80,
      singleGuide('adj', '-4653'),
    );
    const upperBound = getShapePath(
      'smileyFace',
      120,
      80,
      singleGuide('adj', '4653'),
    );

    expect(
      getShapePath('smileyFace', 120, 80, singleGuide('adj', '-10000')),
    ).toBe(lowerBound);
    expect(
      getShapePath('smileyFace', 120, 80, singleGuide('adj', '10000')),
    ).toBe(upperBound);
  });

  it('clamps rounded wedge callout corners to their authored bounds', () => {
    const roundedWedge = (cornerAdjustment: number) =>
      getShapePath(
        'wedgeRoundRectCallout',
        120,
        80,
        guides([
          ['adj1', 'val -20833'],
          ['adj2', 'val 62500'],
          ['adj3', `val ${cornerAdjustment}`],
        ]),
      );
    const lowerBound = roundedWedge(0);
    const upperBound = roundedWedge(50000);

    expect(roundedWedge(-10000)).toBe(lowerBound);
    expect(roundedWedge(100000)).toBe(upperBound);
    expectFinitePath(lowerBound);
    expectFinitePath(upperBound);
  });

  it.each([
    ['right-lower horizontal', 60000, 20000, '8fe9445b', 'ce826378'],
    ['right-upper horizontal', 60000, -20000, '2008646a', 'b4a869ab'],
    ['left-lower horizontal', -60000, 20000, '27092998', 'b2fee451'],
    ['left-upper horizontal', -60000, -20000, '5101aba3', '0bcd7238'],
    ['right-lower vertical', 20000, 60000, 'd953982e', '9b03c237'],
    ['right-upper vertical', 20000, -60000, '64860eef', '4c41013a'],
    ['left-lower vertical', -20000, 60000, 'c5ccdfae', 'ea89ae61'],
    ['left-upper vertical', -20000, -60000, '332cb849', '889bd45a'],
    ['vertical axis', 0, 60000, '2cc46005', '0906c078'],
    ['horizontal axis', 60000, 0, '5739ae7c', 'eb9ad8d5'],
    ['diagonal boundary', 50000, 50000, 'f913fd15', 'aab0e586'],
  ] as const)(
    'routes wedge callouts toward the %s',
    (_route, horizontalAdjustment, verticalAdjustment, rectHash, roundHash) => {
      const adjustmentNode = guides([
        ['adj1', `val ${horizontalAdjustment}`],
        ['adj2', `val ${verticalAdjustment}`],
      ]);
      const paths = {
        rect: hashPath(
          getShapePath('wedgeRectCallout', 120, 80, adjustmentNode),
        ),
        round: hashPath(
          getShapePath('wedgeRoundRectCallout', 120, 80, adjustmentNode),
        ),
      };

      expect(paths).toEqual({ rect: rectHash, round: roundHash });
    },
  );

  it.each([
    'accentBorderCallout1',
    'accentBorderCallout2',
    'accentBorderCallout3',
    'borderCallout1',
    'borderCallout2',
    'borderCallout3',
    'accentCallout1',
    'accentCallout2',
    'accentCallout3',
    'callout1',
    'callout2',
    'callout3',
  ])('ignores an unrelated guide for %s', (shapeType) => {
    expect(getShapePath(shapeType, 120, 80, singleGuide('adj', '25000'))).toBe(
      getShapePath(shapeType, 120, 80, xml({})),
    );
  });

  it('clamps left-right ribbon adjustments to their coupled bounds', () => {
    const ribbon = (
      width: number,
      height: number,
      adjustments: ReadonlyArray<readonly [string, string]>,
    ): string =>
      getShapePath(
        'leftRightRibbon',
        width,
        height,
        guides(
          adjustments.map(([name, value]) => [name, `val ${value}`] as const),
        ),
      );

    expect(ribbon(120, 80, [['adj3', '-10000']])).toBe(
      ribbon(120, 80, [['adj3', '0']]),
    );
    expect(ribbon(120, 80, [['adj3', '100000']])).toBe(
      ribbon(120, 80, [['adj3', '33333']]),
    );
    expect(ribbon(120, 80, [['adj1', '-10000']])).toBe(
      ribbon(120, 80, [['adj1', '0']]),
    );
    expect(
      ribbon(120, 80, [
        ['adj1', '100000'],
        ['adj3', '30000'],
      ]),
    ).toBe(
      ribbon(120, 80, [
        ['adj1', '70000'],
        ['adj3', '30000'],
      ]),
    );
    expect(ribbon(80, 120, [['adj2', '-10000']])).toBe(
      ribbon(80, 120, [['adj2', '0']]),
    );
    expect(ribbon(80, 120, [['adj2', '100000']])).toBe(
      ribbon(80, 120, [['adj2', '46875']]),
    );
  });

  it.each(['ribbon', 'ribbon2'])(
    'clamps PowerPoint ribbon adjustments for %s',
    (shapeType) => {
      const ribbon = (name: string, value: string): string =>
        getShapePath(shapeType, 120, 80, singleGuide(name, value));

      expect(ribbon('adj1', '-10000')).toBe(ribbon('adj1', '0'));
      expect(ribbon('adj1', '100000')).toBe(ribbon('adj1', '33333'));
      expect(ribbon('adj2', '0')).toBe(ribbon('adj2', '25000'));
      expect(ribbon('adj2', '100000')).toBe(ribbon('adj2', '75000'));
    },
  );

  it('keeps common default paths deterministic', () => {
    const rectangle = 'M 0 0 L 120 0 L 120 80 L 0 80 Z';
    expect(getShapePath('rect', 120, 80, xml({}))).toBe(rectangle);
    expect(getShapePath('actionButtonBlank', 120, 80, xml({}))).toBe(rectangle);
    expect(getShapePath('unknown-preset', 120, 80, xml({}))).toBe(rectangle);
    expect(getShapePath('rtTriangle', 120, 80, xml({}))).toBe(
      'M 0 0 L 0 80 L 120 80 Z',
    );
    expect(getShapePath('bentConnector2', 120, 80, xml({}))).toBe(
      'M 120 0 L 120 80 L 0 80',
    );
  });

  it('bounds a safe but extreme authored arc angle', () => {
    const path = getShapePath(
      'chord',
      120,
      80,
      guides([
        ['adj1', `val ${Number.MAX_SAFE_INTEGER}`],
        ['adj2', `val ${-Number.MAX_SAFE_INTEGER}`],
      ]),
    );

    expectFinitePath(path);
    expect(path.length).toBeLessThan(50_000);
  });
});
