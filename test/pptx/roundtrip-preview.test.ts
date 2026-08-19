import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { parse } from '../../src/formats/pptx/parser';
import type {
  Chart,
  CommonChart,
  Group,
  PptxDocument,
  Table,
} from '../../src/formats/pptx/types';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import { validatePptxScene } from '../../src/formats/pptx/scene-validation';
import {
  createPowerPointRoundTripPreview,
  plainTextFromPowerPointHtml,
} from '../../src/formats/pptx/roundtrip/preview';

function sourceScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: {
                flipHorizontal: true,
                flipVertical: false,
                height: 80,
                rotation: 15,
                width: 300,
                x: 20,
                y: 30,
              },
            },
            key: 'source-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'source-run', text: 'Preview text', type: 'run' },
                  ],
                  key: 'source-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'source-slide',
      },
      { elements: [], key: 'empty-slide' },
    ],
    themes: [],
  };
}

function parsedChart(
  chartType: CommonChart['chartType'] = 'barChart',
): CommonChart {
  return {
    barDir: 'col',
    chartType,
    colors: ['#4F46E5'],
    data: [
      {
        key: 'Revenue',
        values: [
          { x: '0', y: 12 },
          { x: '1', y: 18 },
        ],
        xlabels: { '0': 'Q1', '1': 'Q2' },
      },
    ],
    grouping: 'clustered',
    height: 220,
    id: '2',
    left: 40,
    order: 0,
    top: 60,
    type: 'chart',
    width: 420,
  };
}

function chartPreview(element: Chart) {
  const document = {
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [element],
        fill: { type: 'color', value: '#ffffff' },
        layoutElements: [],
        note: '',
      },
    ],
    themeColors: [],
    usedFonts: [],
  } as PptxDocument;
  return createPowerPointRoundTripPreview(document).slides[0]?.elements[0];
}

describe('PowerPoint round-trip semantic preview', () => {
  it('creates a deterministic preservation-only scene from parsed output', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, {
      errorMode: 'strict',
      imageMode: 'none',
    });

    const first = createPowerPointRoundTripPreview(parsed);
    const second = createPowerPointRoundTripPreview(parsed);

    expect(second).toEqual(first);
    expect(validatePptxScene(first)).toEqual({ issues: [], valid: true });
    expect(first.size).toEqual({ height: 540, width: 960 });
    expect(first.slides.map((slide) => slide.key)).toEqual([
      'slide-1',
      'slide-2',
    ]);
    expect(first.slides[1]?.elements).toEqual([]);
    expect(first.slides[0]?.elements).toEqual([
      {
        authored: {},
        key: 'slide-1-element-1',
        name: 'Text Box 2',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: true,
            flipVertical: false,
            height: 80,
            rotation: 15,
            width: 300,
            x: 20,
            y: 30,
          },
        },
        text: {
          body: {
            anchor: 'top',
            vertical: false,
            wrap: true,
          },
          paragraphs: [
            {
              children: [
                {
                  key: 'slide-1-element-1-run-1',
                  text: 'Preview text',
                  type: 'run',
                },
              ],
              key: 'slide-1-element-1-paragraph-1',
            },
          ],
        },
        type: 'text',
      },
    ]);
  });

  it('decodes portable PowerPoint HTML into ordered plain text', () => {
    expect(
      plainTextFromPowerPointHtml(
        '<p><span>A&nbsp;&lt;&amp;</span><br><span>B</span></p><p>C&#x21;</p>',
      ),
    ).toBe('A <&\nB\nC!');
    expect(plainTextFromPowerPointHtml('<p>A</p><p></p><p></p>')).toBe('A');
  });

  it.each([
    ['down', 'bottom'],
    ['mid', 'center'],
    ['dist', 'distributed'],
    ['just', 'justified'],
    ['unknown', 'top'],
  ])('maps vertical alignment %s to %s', async (vAlign, anchor) => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const element = parsed.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected text element');
    element.vAlign = vAlign;

    const preview = createPowerPointRoundTripPreview(parsed);

    expect(preview.slides[0]?.elements[0]).toMatchObject({
      text: { body: { anchor } },
    });
  });

  it('includes a defined autofit mode and omits an absent one', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const element = parsed.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected text element');
    element.autoFit = { type: 'shape' };
    expect(
      createPowerPointRoundTripPreview(parsed).slides[0]?.elements[0],
    ).toMatchObject({ text: { body: { autoFit: 'shape' } } });
    delete element.autoFit;
    const previewElement =
      createPowerPointRoundTripPreview(parsed).slides[0]?.elements[0];
    if (previewElement?.type !== 'text')
      throw new Error('Expected preview text');
    expect(previewElement.text.body).not.toHaveProperty('autoFit');
  });

  it('maps empty native shapes and preserves text-bearing shapes as opaque', () => {
    const nativeShape = {
      content: '',
      height: 25,
      isFlipH: true,
      isFlipV: false,
      left: 5,
      name: 'Native shape',
      rotate: 10,
      top: 6,
      type: 'shape',
      width: 35,
    };
    const unsupportedWithText = {
      content: 'Fallback',
      height: 40,
      isFlipH: false,
      isFlipV: true,
      left: 10,
      name: 'Shape',
      rotate: 5,
      top: 20,
      type: 'shape',
      width: 30,
    };
    const unsupportedWithoutText = {
      height: 20,
      left: 1,
      name: 'Broken',
      top: 2,
      type: 'image',
      width: 30,
    };
    const document = {
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
        {
          elements: [nativeShape, unsupportedWithText, unsupportedWithoutText],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
      ],
      themeColors: [],
      usedFonts: [],
    } as unknown as PptxDocument;

    const preview = createPowerPointRoundTripPreview(document);

    expect(preview.slides[1]?.elements).toEqual([
      {
        authored: {},
        key: 'slide-2-element-1',
        name: 'Native shape',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: true,
            flipVertical: false,
            height: 25,
            rotation: 10,
            width: 35,
            x: 5,
            y: 6,
          },
        },
        type: 'shape',
      },
      {
        authored: {},
        feature: 'shape',
        key: 'slide-2-element-2',
        previewText: 'Fallback',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: false,
            flipVertical: true,
            height: 40,
            rotation: 5,
            width: 30,
            x: 10,
            y: 20,
          },
        },
        type: 'unsupported',
      },
      {
        authored: {},
        key: 'slide-2-element-3',
        resolved: {
          hidden: false,
          transform: { height: 20, width: 30, x: 1, y: 2 },
        },
        type: 'image',
      },
    ]);
  });

  it('does not mutate the parsed document', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const before = structuredClone(parsed);

    createPowerPointRoundTripPreview(parsed);

    expect(parsed).toEqual(before);
  });

  it('maps safe native tables and preserves unsafe grids as opaque', () => {
    const nativeTable: Table = {
      borders: {},
      colWidths: [100, 200],
      data: [
        [
          {
            borders: {
              bottom: {
                borderColor: '#0F172A',
                borderType: 'solid',
                borderWidth: 2,
              },
              left: {
                borderColor: '#F97316',
                borderType: 'dotted',
                borderWidth: 3,
              },
              right: {
                borderColor: '#22C55E',
                borderType: 'solid',
                borderWidth: 4,
              },
              top: {
                borderColor: '#334155',
                borderType: 'dashed',
                borderWidth: 1,
              },
            },
            colSpan: 2,
            fillColor: '#E0F2FE',
            fontBold: true,
            fontColor: '#0F172A',
            rowSpan: 2,
            text: '<b>Alpha &amp; Beta</b>',
            vAlign: 'mid',
          },
          { borders: {}, hMerge: 1, text: 'Value', vAlign: 'down' },
        ],
        [
          {
            borders: {
              top: {
                borderColor: '#000000',
                borderType: 'solid',
                borderWidth: 0,
              },
            },
            text: 'Third',
            vAlign: 'up',
            vMerge: 1,
          },
          {
            borders: {},
            fontBold: false,
            hMerge: 1,
            text: 'Fourth',
            vAlign: 'mid',
            vMerge: 1,
          },
        ],
      ],
      height: 100,
      id: '2',
      isFlipH: false,
      isFlipV: false,
      left: 10,
      name: 'Native table',
      order: 0,
      rotate: 0,
      rowHeights: [40, 60],
      top: 20,
      type: 'table',
      width: 300,
    };
    const noNameTable = structuredClone(nativeTable);
    noNameTable.id = '3';
    delete noNameTable.name;
    const invalidTables = [
      (() => {
        const value = structuredClone(nativeTable);
        value.colWidths = [];
        return value;
      })(),
      (() => {
        const value = structuredClone(nativeTable);
        value.data = [];
        return value;
      })(),
      (() => {
        const value = structuredClone(nativeTable);
        value.colWidths = [0, 300];
        return value;
      })(),
      (() => {
        const value = structuredClone(nativeTable);
        value.rowHeights = [0, 100];
        return value;
      })(),
      (() => {
        const value = structuredClone(nativeTable);
        value.rowHeights = [100];
        return value;
      })(),
      (() => {
        const value = structuredClone(nativeTable);
        value.data[1] = [value.data[1]?.[0] as Table['data'][number][number]];
        return value;
      })(),
      { ...structuredClone(nativeTable), width: 301 },
      { ...structuredClone(nativeTable), height: 101 },
      { ...structuredClone(nativeTable), left: Number.NaN },
    ];
    const document = {
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [nativeTable, noNameTable, ...invalidTables],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
      ],
      themeColors: [],
      usedFonts: [],
    } as unknown as PptxDocument;

    const preview = createPowerPointRoundTripPreview(document);

    expect(preview.slides[0]?.elements[0]).toEqual({
      authored: {},
      columns: [100, 200],
      key: 'slide-1-element-1',
      name: 'Native table',
      resolved: {
        hidden: false,
        transform: {
          flipHorizontal: false,
          flipVertical: false,
          height: 100,
          rotation: 0,
          width: 300,
          x: 10,
          y: 20,
        },
      },
      rows: [
        {
          cells: [
            {
              borders: {
                bottom: { color: '#0F172A', style: 'solid', width: 2 },
                left: { color: '#F97316', style: 'dotted', width: 3 },
                right: { color: '#22C55E', style: 'solid', width: 4 },
                top: { color: '#334155', style: 'dashed', width: 1 },
              },
              colSpan: 2,
              fillColor: '#E0F2FE',
              rowSpan: 2,
              text: {
                body: { anchor: 'center' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-1-cell-1-run-1',
                        properties: { bold: true, color: '#0F172A' },
                        text: 'Alpha & Beta',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-1-cell-1-paragraph-1',
                  },
                ],
              },
            },
            {
              borders: {},
              hMerge: true,
              text: {
                body: { anchor: 'bottom' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-1-cell-2-run-1',
                        text: 'Value',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-1-cell-2-paragraph-1',
                  },
                ],
              },
            },
          ],
          height: 40,
        },
        {
          cells: [
            {
              borders: {},
              text: {
                body: { anchor: 'top' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-2-cell-1-run-1',
                        text: 'Third',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-2-cell-1-paragraph-1',
                  },
                ],
              },
              vMerge: true,
            },
            {
              borders: {},
              hMerge: true,
              text: {
                body: { anchor: 'center' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-2-cell-2-run-1',
                        properties: { bold: false },
                        text: 'Fourth',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-2-cell-2-paragraph-1',
                  },
                ],
              },
              vMerge: true,
            },
          ],
          height: 60,
        },
      ],
      type: 'table',
    });
    expect(preview.slides[0]?.elements[1]).not.toHaveProperty('name');
    expect(preview.slides[0]?.elements.slice(2)).toHaveLength(
      invalidTables.length,
    );
    expect(
      preview.slides[0]?.elements.slice(2).map((element) => element.type),
    ).toEqual(invalidTables.map(() => 'unsupported'));
    expect(validatePptxScene(preview)).toEqual({ issues: [], valid: true });
  });

  it('dispatches native groups with stable nested semantic keys', () => {
    const group: Group = {
      childSpace: { height: 50, width: 100, x: 10, y: 20 },
      elements: [
        {
          borderColor: '#000000',
          borderStrokeDasharray: '0',
          borderType: 'solid',
          borderWidth: 0,
          content: '',
          fill: null,
          height: 20,
          id: '3',
          isFlipH: false,
          isFlipV: false,
          left: 30,
          name: 'Child',
          order: 0,
          rotate: 0,
          shapType: 'rect',
          top: 40,
          type: 'shape',
          vAlign: 'up',
          width: 25,
          wrap: true,
        },
      ],
      height: 100,
      id: '2',
      isFlipH: false,
      isFlipV: false,
      left: 20,
      order: 0,
      rotate: 0,
      top: 30,
      type: 'group',
      width: 200,
    };
    const document = {
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [group],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
      ],
      themeColors: [],
      usedFonts: [],
    } as PptxDocument;

    expect(
      createPowerPointRoundTripPreview(document).slides[0]?.elements[0],
    ).toMatchObject({
      elements: [
        {
          key: 'slide-1-element-1-element-1',
          type: 'shape',
        },
      ],
      key: 'slide-1-element-1',
      resolved: {
        transform: {
          childSpace: { height: 50, width: 100, x: 10, y: 20 },
        },
      },
      type: 'group',
    });
  });

  it.each(['barChart', 'lineChart', 'pieChart', 'doughnutChart'] as const)(
    'maps a safe native %s with stable series keys',
    (chartType) => {
      const element = parsedChart(chartType);
      if (chartType === 'lineChart') element.marker = true;
      if (chartType === 'doughnutChart') element.holeSize = '55';

      expect(chartPreview(element)).toMatchObject({
        barDirection: 'col',
        chartType,
        grouping: 'clustered',
        key: 'slide-1-element-1',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: false,
            flipVertical: false,
            height: 220,
            rotation: 0,
            width: 420,
            x: 40,
            y: 60,
          },
        },
        series: [
          {
            categories: ['Q1', 'Q2'],
            color: '#4F46E5',
            key: 'slide-1-element-1-series-1',
            name: 'Revenue',
            values: [12, 18],
          },
        ],
        type: 'chart',
        ...(chartType === 'lineChart' ? { marker: true } : {}),
        ...(chartType === 'doughnutChart' ? { holeSize: 55 } : {}),
      });
    },
  );

  it('omits absent optional chart fields and invalid colors', () => {
    for (const color of ['invalid', 'x#4F46E5', '#4F46E5x']) {
      const element = parsedChart('lineChart');
      delete element.barDir;
      delete element.grouping;
      element.colors = [color];

      const preview = chartPreview(element);
      if (preview?.type !== 'chart') throw new Error('Expected native chart');
      expect(preview).not.toHaveProperty('barDirection');
      expect(preview).not.toHaveProperty('grouping');
      expect(preview).not.toHaveProperty('marker');
      expect(preview.series[0]).not.toHaveProperty('color');
    }
  });

  it('retains every allowed chart grouping value', () => {
    for (const grouping of [
      'clustered',
      'percentStacked',
      'stacked',
      'standard',
    ] as const) {
      const element = parsedChart('lineChart');
      element.grouping = grouping;
      expect(chartPreview(element)).toMatchObject({ grouping, type: 'chart' });
    }
  });

  it.each<[string, (value: CommonChart) => void]>([
    [
      'scatter type',
      (value) => {
        value.chartType = 'scatterChart' as never;
      },
    ],
    [
      'array series',
      (value) => {
        value.data = [[1, 2]] as never;
      },
    ],
    [
      'non-string key',
      (value) => {
        value.data[0]!.key = 7 as never;
      },
    ],
    [
      'blank key',
      (value) => {
        value.data[0]!.key = '  ';
      },
    ],
    [
      'non-array values',
      (value) => {
        value.data[0]!.values = null as never;
      },
    ],
    [
      'empty values',
      (value) => {
        value.data[0]!.values = [];
      },
    ],
    [
      'non-string x',
      (value) => {
        value.data[0]!.values[0]!.x = 0 as never;
      },
    ],
    [
      'non-number y',
      (value) => {
        value.data[0]!.values[0]!.y = '12' as never;
      },
    ],
    [
      'non-finite y',
      (value) => {
        value.data[0]!.values[0]!.y = Number.NaN;
      },
    ],
    [
      'invalid grouping',
      (value) => {
        value.grouping = 'invalid';
      },
    ],
    [
      'fractional hole',
      (value: CommonChart) => {
        value.chartType = 'doughnutChart';
        value.holeSize = '50.5';
      },
    ],
    [
      'small hole',
      (value: CommonChart) => {
        value.chartType = 'doughnutChart';
        value.holeSize = '9';
      },
    ],
    [
      'large hole',
      (value: CommonChart) => {
        value.chartType = 'doughnutChart';
        value.holeSize = '91';
      },
    ],
  ])('preserves an unsafe %s as opaque', (_name, mutate) => {
    const element = parsedChart();
    mutate(element);

    expect(chartPreview(element)).toMatchObject({
      feature: 'chart',
      type: 'unsupported',
    });
  });

  it('requires exactly one pie-family series but permits multiple bar series', () => {
    for (const chartType of ['pieChart', 'doughnutChart'] as const) {
      const element = parsedChart(chartType);
      element.data.push(structuredClone(element.data[0]!));
      expect(chartPreview(element)).toMatchObject({
        feature: 'chart',
        type: 'unsupported',
      });
    }
    const bar = parsedChart();
    bar.data.push({
      key: 'Costs',
      values: [{ x: '0', y: 5 }],
      xlabels: { '0': 'Q1' },
    });
    expect(chartPreview(bar)).toMatchObject({
      series: [{ name: 'Revenue' }, { name: 'Costs' }],
      type: 'chart',
    });
  });

  it.each([10, 90])('retains chart hole-size boundary %d', (holeSize) => {
    const element = parsedChart('doughnutChart');
    element.holeSize = String(holeSize);
    expect(chartPreview(element)).toMatchObject({ holeSize, type: 'chart' });
  });

  it.each([
    ['left', Number.NaN],
    ['top', Number.POSITIVE_INFINITY],
    ['width', Number.NaN],
    ['width', 0],
    ['height', Number.NEGATIVE_INFINITY],
    ['height', 0],
  ] as const)(
    'omits a transform with non-rendering %s %s from the portable preview',
    async (property, value) => {
      const source = await createPptx(sourceScene());
      const parsed = await parse(source.data, { imageMode: 'none' });
      const element = parsed.slides[0]?.elements[0];
      if (!element) throw new Error('Expected a parsed preview element');
      element[property] = value;

      const preview = createPowerPointRoundTripPreview(parsed);

      expect(preview.slides[0]?.elements[0]?.resolved).toEqual({
        hidden: false,
      });
      expect(validatePptxScene(preview)).toEqual({ issues: [], valid: true });
    },
  );
});
