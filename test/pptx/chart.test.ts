import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getChartInfo } from '../../src/formats/pptx/internal/chart';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import type { ChartType } from '../../src/formats/pptx/types';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(): PptxParserContext {
  return {
    themeContent: xml({
      'a:theme': {
        'a:themeElements': {
          'a:clrScheme': {
            'a:accent1': { 'a:srgbClr': { attrs: { val: '204060' } } },
          },
        },
      },
    }),
  } as unknown as PptxParserContext;
}

function commonSeries(extra: object = {}): XmlLookupValue {
  return xml({
    'c:tx': {
      'c:strRef': {
        'c:strCache': {
          'c:pt': [
            { attrs: { idx: '0' }, 'c:v': 'Revenue' },
            { attrs: { idx: '1' }, 'c:v': 'ignored' },
          ],
        },
      },
    },
    'c:cat': {
      'c:strRef': {
        'c:strCache': {
          'c:pt': [
            { attrs: { idx: '0' }, 'c:v': 'Q1' },
            { attrs: { idx: '1' }, 'c:v': 'Q2' },
          ],
        },
      },
    },
    'c:val': {
      'c:numRef': {
        'c:numCache': {
          'c:pt': [
            { attrs: { idx: '0' }, 'c:v': '12.5' },
            { attrs: { idx: '1' }, 'c:v': '-3' },
          ],
        },
      },
    },
    'c:spPr': {
      'a:solidFill': { 'a:srgbClr': { attrs: { val: 'abcdef' } } },
    },
    ...extra,
  });
}

const COMMON_TYPES: readonly ChartType[] = [
  'lineChart',
  'line3DChart',
  'barChart',
  'bar3DChart',
  'pieChart',
  'pie3DChart',
  'doughnutChart',
  'areaChart',
  'area3DChart',
  'radarChart',
  'surfaceChart',
  'surface3DChart',
  'stockChart',
];

describe('PPTX chart parsing', () => {
  it.each(COMMON_TYPES)('recognizes %s', (type) => {
    const result = getChartInfo(
      xml({ [`c:${type}`]: { 'c:ser': commonSeries() } }),
      context(),
    );

    expect(result?.type).toBe(type);
  });

  it('returns null when no supported chart series exists', () => {
    expect(getChartInfo(xml({}), context())).toBeNull();
    expect(
      getChartInfo(
        xml({ 'c:unknownChart': { 'c:ser': commonSeries() } }),
        context(),
      ),
    ).toBeNull();
    expect(
      getChartInfo(xml({ 'c:barChart': { unrelated: {} } }), context()),
    ).toBeNull();
  });

  it('extracts series names, values, category labels, and direct colors', () => {
    const result = getChartInfo(
      xml({ 'c:barChart': { 'c:ser': commonSeries() } }),
      context(),
    );

    expect(result).toMatchObject({
      colors: ['#abcdef'],
      data: [
        {
          key: 'Revenue',
          values: [
            { x: '0', y: 12.5 },
            { x: '1', y: -3 },
          ],
          xlabels: { '0': 'Q1', '1': 'Q2' },
        },
      ],
      type: 'barChart',
    });
  });

  it('supports unprefixed chart keys and numeric category caches', () => {
    const series = commonSeries({
      'c:cat': {
        'c:numRef': {
          'c:numCache': {
            'c:pt': { attrs: { idx: '4' }, 'c:v': '2026' },
          },
        },
      },
      'c:val': {
        'c:numRef': {
          'c:numCache': { 'c:pt': { 'c:v': '7' } },
        },
      },
    });

    expect(
      getChartInfo(xml({ barChart: { 'c:ser': series } }), context()),
    ).toMatchObject({
      data: [
        {
          values: [{ x: '', y: 7 }],
          xlabels: { '4': '2026' },
        },
      ],
      type: 'barChart',
    });
  });

  it.each(['12x', 'not-a-number', 'Infinity', '-Infinity'])(
    'normalizes malformed chart value %j to zero',
    (value) => {
      const series = commonSeries({
        'c:val': {
          'c:numRef': {
            'c:numCache': {
              'c:pt': { attrs: { idx: '7' }, 'c:v': value },
            },
          },
        },
      });
      const result = getChartInfo(
        xml({ 'c:barChart': { 'c:ser': series } }),
        context(),
      );

      expect(result?.data).toEqual([
        {
          key: 'Revenue',
          values: [{ x: '7', y: 0 }],
          xlabels: { '0': 'Q1', '1': 'Q2' },
        },
      ]);
    },
  );

  it('extracts scatter and bubble axes without non-finite values', () => {
    const scatterSeries = (x: string, y: string) =>
      xml({
        'c:xVal': {
          'c:numRef': { 'c:numCache': { 'c:pt': { 'c:v': x } } },
        },
        'c:yVal': {
          'c:numRef': { 'c:numCache': { 'c:pt': { 'c:v': y } } },
        },
      });
    const series = [
      scatterSeries('1.5', '2.5'),
      scatterSeries('unused', 'Infinity'),
    ];

    expect(
      getChartInfo(xml({ 'c:scatterChart': { 'c:ser': series } }), context()),
    ).toMatchObject({
      colors: ['', ''],
      data: [[1.5], [2.5], [0]],
      type: 'scatterChart',
    });
    expect(
      getChartInfo(xml({ 'c:bubbleChart': { 'c:ser': series } }), context())
        ?.type,
    ).toBe('bubbleChart');
  });

  it('extracts chart options and validates bar direction', () => {
    const result = getChartInfo(
      xml({
        'c:lineChart': {
          'c:ser': commonSeries(),
          'c:barDir': { attrs: { val: 'col' } },
          'c:grouping': { attrs: { val: 'stacked' } },
          'c:holeSize': { attrs: { val: '60' } },
          'c:marker': {},
          'c:radarStyle': { attrs: { val: 'filled' } },
        },
      }),
      context(),
    );

    expect(result).toMatchObject({
      barDir: 'col',
      grouping: 'stacked',
      holeSize: '60',
      marker: true,
      style: 'filled',
    });
    expect(
      getChartInfo(
        xml({
          'c:barChart': {
            'c:barDir': { attrs: { val: 'diagonal' } },
            'c:ser': commonSeries(),
          },
        }),
        context(),
      ),
    ).not.toHaveProperty('barDir');

    const horizontal = getChartInfo(
      xml({
        'c:barChart': {
          'c:barDir': { attrs: { val: 'bar' } },
          'c:marker': {},
          'c:ser': commonSeries(),
        },
      }),
      context(),
    );
    expect(horizontal).toMatchObject({ barDir: 'bar' });
    expect(horizontal).not.toHaveProperty('marker');
  });

  it('resolves scheme colors with tint and preserves stock color semantics', () => {
    const schemeSeries = commonSeries({
      'c:spPr': {
        'a:solidFill': {
          'a:schemeClr': {
            attrs: { val: 'accent1' },
            'a:tint': { attrs: { val: '50000' } },
          },
        },
      },
    });
    expect(
      getChartInfo(xml({ 'c:barChart': { 'c:ser': schemeSeries } }), context())
        ?.colors,
    ).toEqual(['#709fcf']);
    expect(
      getChartInfo(
        xml({ 'c:stockChart': { 'c:ser': schemeSeries } }),
        context(),
      )?.colors,
    ).toEqual([]);
  });

  it.each(['pieChart', 'pie3DChart', 'doughnutChart'] as const)(
    'uses data-point colors for %s',
    (type) => {
      const series = commonSeries({
        'c:dPt': [
          {
            'c:spPr': {
              'a:solidFill': {
                'a:srgbClr': { attrs: { val: 'ff0000' } },
              },
            },
          },
          {
            'c:spPr': {
              'a:solidFill': {
                'a:srgbClr': { attrs: { val: '00ff00' } },
              },
            },
          },
        ],
      });

      expect(
        getChartInfo(xml({ [`c:${type}`]: { 'c:ser': series } }), context())
          ?.colors,
      ).toEqual(['#ff0000', '#00ff00']);
    },
  );

  it('resolves scheme colors from line and marker outlines', () => {
    const scheme = { 'a:schemeClr': { attrs: { val: 'accent1' } } };
    const lineSeries = commonSeries({
      'c:spPr': { 'a:ln': { 'a:solidFill': scheme } },
    });
    const markerSeries = commonSeries({
      'c:marker': {
        'c:spPr': { 'a:ln': { 'a:solidFill': scheme } },
      },
      'c:spPr': {},
    });

    expect(
      getChartInfo(
        xml({ 'c:lineChart': { 'c:ser': [lineSeries, markerSeries] } }),
        context(),
      )?.colors,
    ).toEqual(['#204060', '#204060']);
  });

  it('normalizes invalid authored colors and reads scatter style', () => {
    const invalidColorSeries = commonSeries({
      'c:spPr': {
        'a:solidFill': { 'a:srgbClr': { attrs: { val: 'invalid' } } },
      },
    });
    const result = getChartInfo(
      xml({
        'c:scatterChart': {
          'c:scatterStyle': { attrs: { val: 'smoothMarker' } },
          'c:ser': invalidColorSeries,
        },
      }),
      context(),
    );

    expect(result).toMatchObject({ colors: [''], style: 'smoothMarker' });
  });
});
