import { describe, expect, it } from 'vitest';

import {
  MAX_POWERPOINT_CREATION_CHART_POINTS,
  MAX_POWERPOINT_CREATION_CHART_SERIES,
} from '../../src/formats/pptx/creation-limits';
import type {
  PptxSceneChartElement,
  PptxSceneChartType,
  PptxSceneDocument,
} from '../../src/formats/pptx/scene-types';
import { validatePptxScene } from '../../src/formats/pptx/scene-validation';

function chart(
  chartType: PptxSceneChartType = 'barChart',
): PptxSceneChartElement {
  return {
    authored: {
      transform: { height: 240, width: 480, x: 40, y: 60 },
    },
    chartType,
    key: 'chart-1',
    resolved: { hidden: false },
    series: [
      {
        categories: ['Q1', 'Q2', 'Q3'],
        color: '#4F46E5',
        key: 'series-1',
        name: 'Revenue',
        values: [10, 20, 30],
      },
    ],
    type: 'chart',
  };
}

function scene(element: PptxSceneChartElement = chart()): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [{ elements: [element], key: 'slide-1' }],
    themes: [],
  };
}

function chartIssues(
  element: PptxSceneChartElement,
  profile: 'create-native-v1' | 'create-text-v1' = 'create-native-v1',
) {
  return validatePptxScene(scene(element), { profile }).issues;
}

describe('native PowerPoint chart scene validation', () => {
  it.each<PptxSceneChartType>([
    'barChart',
    'doughnutChart',
    'lineChart',
    'pieChart',
  ])('accepts a bounded %s scene', (chartType) => {
    const element = chart(chartType);
    if (chartType === 'barChart') {
      element.barDirection = 'col';
      element.grouping = 'clustered';
    } else if (chartType === 'lineChart') {
      element.grouping = 'standard';
      element.marker = true;
    } else if (chartType === 'doughnutChart') {
      element.holeSize = 50;
    }

    expect(
      validatePptxScene(scene(element), { profile: 'create-native-v1' }),
    ).toEqual({ issues: [], valid: true });
  });

  it('keeps native charts outside the text-only creation profile', () => {
    expect(
      validatePptxScene(scene(), { profile: 'create-text-v1' }).issues,
    ).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });
  });

  it('requires one bounded non-empty series with aligned data', () => {
    const element = chart();
    element.series[0]!.categories = ['Q1'];
    element.series[0]!.values = [1, Number.NaN];
    const result = validatePptxScene(scene(element), {
      profile: 'create-native-v1',
    });

    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          path: '$.slides[0].elements[0].series[0]',
        }),
        expect.objectContaining({
          code: 'invalid-numeric-value',
          path: '$.slides[0].elements[0].series[0].values[1]',
        }),
      ]),
    );
  });

  it('enforces chart series and point budgets', () => {
    const tooManySeries = chart();
    tooManySeries.series = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_SERIES + 1 },
      (_, index) => ({
        categories: ['A'],
        key: `series-${index}`,
        name: `Series ${index}`,
        values: [index],
      }),
    );
    const tooManyPoints = chart();
    tooManyPoints.series[0]!.categories = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_POINTS + 1 },
      () => 'A',
    );
    tooManyPoints.series[0]!.values = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_POINTS + 1 },
      () => 1,
    );

    for (const value of [tooManySeries, tooManyPoints]) {
      expect(
        validatePptxScene(scene(value), { profile: 'create-native-v1' }).issues,
      ).toContainEqual(
        expect.objectContaining({ code: 'resource-limit-exceeded' }),
      );
    }
  });

  it('requires exactly one series for pie-family charts', () => {
    for (const chartType of ['pieChart', 'doughnutChart'] as const) {
      const element = chart(chartType);
      element.series.push({
        ...structuredClone(element.series[0]!),
        key: 's2',
      });
      expect(
        validatePptxScene(scene(element), {
          profile: 'create-native-v1',
        }).issues,
      ).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Pie and doughnut charts require exactly one series',
        path: '$.slides[0].elements[0].series',
      });
    }
  });

  it('rejects options that do not belong to the selected chart type', () => {
    const element = chart('pieChart');
    element.barDirection = 'col';
    element.grouping = 'clustered';
    element.holeSize = 91;
    element.marker = true;

    const result = validatePptxScene(scene(element), {
      profile: 'create-native-v1',
    });
    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'invalid-numeric-value',
          path: '$.slides[0].elements[0].holeSize',
        }),
        {
          code: 'unsupported-feature',
          message: 'Chart options must match the selected native chart type',
          path: '$.slides[0].elements[0]',
        },
      ]),
    );
  });

  it('rejects unsafe names, duplicate keys, and unknown properties', () => {
    const element = chart();
    element.series[0]!.name = 'bad\u0000';
    element.series.push({
      ...structuredClone(element.series[0]!),
      key: 'series-1',
      name: 'Duplicate',
    });
    (element as unknown as Record<string, unknown>).unknown = true;

    const result = validatePptxScene(scene(element));
    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'invalid-office-text-escape',
          path: '$.slides[0].elements[0].series[0].name',
        }),
        expect.objectContaining({
          code: 'duplicate-public-key',
          path: '$.slides[0].elements[0].series[1].key',
        }),
        {
          code: 'invalid-scene-document',
          message: 'Unknown property',
          path: '$.slides[0].elements[0].unknown',
        },
      ]),
    );
  });

  it('reports an empty chart series list exactly', () => {
    const element = chart();
    element.series = [];
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message: 'A chart needs at least one series',
      path: '$.slides[0].elements[0].series',
    });
  });

  it('accepts exact chart series and point boundaries', () => {
    const seriesBoundary = chart();
    seriesBoundary.series = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_SERIES },
      (_, index) => ({
        categories: ['A'],
        key: `series-${index}`,
        name: `Series ${index}`,
        values: [index],
      }),
    );
    expect(chartIssues(seriesBoundary)).toEqual([]);

    const pointBoundary = chart();
    pointBoundary.series[0]!.categories = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_POINTS },
      () => 'A',
    );
    pointBoundary.series[0]!.values = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_POINTS },
      () => -1,
    );
    expect(chartIssues(pointBoundary)).toEqual([]);
  });

  it('reports the exact chart series budget message', () => {
    const element = chart();
    element.series = Array.from(
      { length: MAX_POWERPOINT_CREATION_CHART_SERIES + 1 },
      (_, index) => ({
        categories: ['A'],
        key: `series-${index}`,
        name: `Series ${index}`,
        values: [index],
      }),
    );
    expect(chartIssues(element)).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `A chart supports at most ${MAX_POWERPOINT_CREATION_CHART_SERIES} series`,
      path: '$.slides[0].elements[0].series',
    });
  });

  it('rejects an invalid chart series color exactly', () => {
    const element = chart();
    element.series[0]!.color = 'invalid';
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a #RRGGBB color',
      path: '$.slides[0].elements[0].series[0].color',
    });
  });

  it.each([
    [7, 'Expected a non-empty chart series name'],
    ['  ', 'Expected a non-empty chart series name'],
    ['bad\u0000', 'Chart series name cannot be serialized safely'],
  ] as const)('rejects chart series name %# exactly', (name, message) => {
    const element = chart();
    element.series[0]!.name = name as never;
    expect(chartIssues(element)).toContainEqual({
      code:
        typeof name === 'string' && name.includes('\u0000')
          ? 'invalid-office-text-escape'
          : 'invalid-scene-document',
      message,
      path: '$.slides[0].elements[0].series[0].name',
    });
  });

  it('accepts a non-empty sentinel chart series name', () => {
    const element = chart();
    element.series[0]!.name = 'Stryker was here!';
    expect(chartIssues(element)).toEqual([]);
  });

  it.each([
    ['categories', null, 'Expected an array'],
    ['values', null, 'Expected an array'],
  ] as const)('rejects non-array chart %s exactly', (key, value, message) => {
    const element = chart();
    element.series[0]![key] = value as never;
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message,
      path: `$.slides[0].elements[0].series[0].${key}`,
    });
  });

  it('reports chart category/value length mismatch exactly', () => {
    const element = chart();
    element.series[0]!.categories = ['A'];
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Chart categories and values must have equal lengths',
      path: '$.slides[0].elements[0].series[0]',
    });
  });

  it('enforces each side of the per-series chart point budget', () => {
    for (const key of ['categories', 'values'] as const) {
      const element = chart();
      element.series[0]!.categories = [];
      element.series[0]!.values = [];
      element.series[0]![key] = Array.from(
        { length: MAX_POWERPOINT_CREATION_CHART_POINTS + 1 },
        () => (key === 'categories' ? 'A' : 1),
      ) as never;
      expect(chartIssues(element)).toContainEqual({
        code: 'resource-limit-exceeded',
        message: `A chart series supports at most ${MAX_POWERPOINT_CREATION_CHART_POINTS} points`,
        path: '$.slides[0].elements[0].series[0]',
      });
    }
  });

  it('validates every chart category and numeric value', () => {
    const category = chart();
    category.series[0]!.categories[1] = 7 as never;
    expect(chartIssues(category)).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected text content',
      path: '$.slides[0].elements[0].series[0].categories[1]',
    });
    const value = chart();
    value.series[0]!.values[1] = Number.NaN;
    expect(chartIssues(value)).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Expected a finite number',
      path: '$.slides[0].elements[0].series[0].values[1]',
    });
  });

  it.each([
    ['chartType', 'unknown', 'Unknown native chart type'],
    ['barDirection', 'unknown', 'Unknown chart bar direction'],
    ['grouping', 'unknown', 'Unknown chart grouping'],
  ] as const)('rejects chart %s exactly', (key, value, message) => {
    const element = chart();
    (element as unknown as Record<string, unknown>)[key] = value;
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message,
      path: `$.slides[0].elements[0].${key}`,
    });
  });

  it('accepts every native bar direction and grouping value', () => {
    for (const barDirection of ['bar', 'col'] as const) {
      const element = chart('barChart');
      element.barDirection = barDirection;
      expect(chartIssues(element)).toEqual([]);
    }
    for (const grouping of [
      'clustered',
      'percentStacked',
      'stacked',
      'standard',
    ] as const) {
      const element = chart('lineChart');
      element.grouping = grouping;
      expect(chartIssues(element)).toEqual([]);
    }
  });

  it('rejects a non-boolean chart marker exactly', () => {
    const element = chart('lineChart');
    element.marker = 'yes' as never;
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a boolean',
      path: '$.slides[0].elements[0].marker',
    });
  });

  it.each([9, 10.5, 91])('rejects chart hole size %s exactly', (holeSize) => {
    const element = chart('doughnutChart');
    element.holeSize = holeSize;
    expect(chartIssues(element)).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Doughnut hole size must be an integer from 10 through 90',
      path: '$.slides[0].elements[0].holeSize',
    });
  });

  it.each([10, 90])('accepts chart hole size boundary %s', (holeSize) => {
    const element = chart('doughnutChart');
    element.holeSize = holeSize;
    expect(chartIssues(element)).toEqual([]);
  });

  it.each([
    ['barDirection', 'col', 'pieChart'],
    ['grouping', 'clustered', 'pieChart'],
    ['holeSize', 50, 'barChart'],
    ['marker', true, 'barChart'],
  ] as const)(
    'rejects incompatible chart option %s independently',
    (key, option, chartType) => {
      const element = chart(chartType);
      (element as unknown as Record<string, unknown>)[key] = option;
      expect(chartIssues(element)).toContainEqual({
        code: 'unsupported-feature',
        message: 'Chart options must match the selected native chart type',
        path: '$.slides[0].elements[0]',
      });
    },
  );

  it.each([
    ['flipHorizontal', true],
    ['flipVertical', true],
    ['rotation', 15],
  ] as const)('rejects native chart transform %s exactly', (key, value) => {
    const element = chart();
    const transform = element.authored.transform!;
    (transform as unknown as Record<string, unknown>)[key] = value;
    expect(chartIssues(element)).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Native chart creation supports unrotated, unflipped graphic frames only',
      path: '$.slides[0].elements[0].authored.transform',
    });
  });

  it('accepts explicit zero rotation and false chart flips', () => {
    const element = chart();
    const transform = element.authored.transform!;
    transform.flipHorizontal = false;
    transform.flipVertical = false;
    transform.rotation = 0;
    expect(chartIssues(element)).toEqual([]);
  });

  it('does not claim native transform support in the text profile', () => {
    const element = chart();
    element.authored.transform!.rotation = 15;
    expect(
      chartIssues(element, 'create-text-v1').filter(
        (issue) =>
          issue.message ===
          'Native chart creation supports unrotated, unflipped graphic frames only',
      ),
    ).toEqual([]);
  });

  it('rejects native chart shape styling and a missing transform exactly', () => {
    const styled = chart();
    styled.authored.fillColor = '#FFFFFF';
    expect(chartIssues(styled)).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 does not apply shape styling to charts',
      path: '$.slides[0].elements[0].authored',
    });
    const missing = chart();
    delete missing.authored.transform;
    expect(chartIssues(missing)).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 requires an authored chart transform',
      path: '$.slides[0].elements[0].authored.transform',
    });
  });
});
