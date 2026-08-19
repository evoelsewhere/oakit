import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  renderPptxToSvg,
  type PptxSceneChartElement,
  type PptxSceneChartType,
  type PptxSceneDocument,
} from '../../src';

function chart(
  chartType: PptxSceneChartType,
  index: number,
): PptxSceneChartElement {
  const element: PptxSceneChartElement = {
    authored: {
      transform: { height: 220, width: 420, x: 40, y: 60 },
    },
    chartType,
    key: `chart-${index}`,
    name: `${chartType} example`,
    resolved: { hidden: false },
    series: [
      {
        categories: ['Q1', 'Q2', 'Q3'],
        color: '#4F46E5',
        key: `chart-${index}-series-1`,
        name: 'Revenue',
        values: [12, 18, 27],
      },
    ],
    type: 'chart',
  };
  if (chartType === 'barChart') {
    element.barDirection = 'col';
    element.grouping = 'clustered';
  } else if (chartType === 'lineChart') {
    element.grouping = 'standard';
    element.marker = true;
  } else if (chartType === 'doughnutChart') {
    element.holeSize = 55;
  }
  return element;
}

function scene(): PptxSceneDocument {
  const chartTypes: PptxSceneChartType[] = [
    'barChart',
    'lineChart',
    'pieChart',
    'doughnutChart',
  ];
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: chartTypes.map((chartType, index) => ({
      elements: [chart(chartType, index + 1)],
      key: `slide-${index + 1}`,
    })),
    themes: [],
  };
}

describe('native PowerPoint chart creation', () => {
  it('creates, strict-parses, and Office-free renders common native charts', async () => {
    const created = await createPptx(scene());
    const [parsed, rendered, archive] = await Promise.all([
      parsePptx(created.data, { errorMode: 'strict', imageMode: 'none' }),
      renderPptxToSvg(created.data),
      JSZip.loadAsync(created.data),
    ]);

    expect(created.report).toMatchObject({
      level: 'C2',
      supportProfile: { id: 'pptx-create-native-v1' },
    });
    expect(parsed.slides.map((slide) => slide.elements[0])).toMatchObject([
      {
        barDir: 'col',
        chartType: 'barChart',
        data: [
          {
            key: 'Revenue',
            values: [
              { x: '0', y: 12 },
              { x: '1', y: 18 },
              { x: '2', y: 27 },
            ],
          },
        ],
        grouping: 'clustered',
        type: 'chart',
      },
      {
        chartType: 'lineChart',
        grouping: 'standard',
        marker: true,
        type: 'chart',
      },
      { chartType: 'pieChart', type: 'chart' },
      { chartType: 'doughnutChart', holeSize: '55', type: 'chart' },
    ]);
    expect(rendered.slides).toHaveLength(4);
    const chartTypes = ['barChart', 'lineChart', 'pieChart', 'doughnutChart'];
    for (const [index, slide] of rendered.slides.entries()) {
      const svg = new TextDecoder().decode(slide.data);
      expect(svg).toContain('<svg');
      expect(svg).toContain(chartTypes[index]);
      expect(slide.warnings).toContainEqual(
        expect.objectContaining({ code: 'approximate-chart' }),
      );
    }
    for (let index = 1; index <= 4; index += 1) {
      expect(archive.file(`ppt/charts/chart${index}.xml`)).not.toBeNull();
      const relationships = await archive
        .file(`ppt/slides/_rels/slide${index}.xml.rels`)
        ?.async('text');
      expect(relationships).toContain(`../charts/chart${index}.xml`);
    }
  });

  it('is byte deterministic across independent chart writes', async () => {
    const [first, second] = await Promise.all([
      createPptx(scene()),
      createPptx(scene()),
    ]);

    expect(second).toEqual(first);
  });
});
