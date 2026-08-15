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
