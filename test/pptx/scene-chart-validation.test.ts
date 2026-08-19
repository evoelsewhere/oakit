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
});
