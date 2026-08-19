import { describe, expect, it } from 'vitest';

import type { PptxSceneChartElement } from '../../src/formats/pptx/scene-types';
import {
  serializeChartFrame,
  serializeChartPart,
} from '../../src/formats/pptx/writer/chart';

function chart(
  chartType: PptxSceneChartElement['chartType'] = 'barChart',
): PptxSceneChartElement {
  return {
    authored: {
      hidden: true,
      transform: { height: 240, width: 480, x: 40, y: 60 },
    },
    chartType,
    description: 'Chart <description>',
    key: 'chart-1',
    name: 'Revenue & growth',
    resolved: { hidden: false },
    series: [
      {
        categories: ['Q1', 'Q2 & Q3'],
        color: '#4f46e5',
        key: 'series-1',
        name: 'Revenue <net>',
        values: [10, -20.5],
      },
    ],
    title: 'Accessible "chart"',
    type: 'chart',
  };
}

describe('native PowerPoint chart serialization', () => {
  it('serializes a deterministic cache-backed bar chart', () => {
    const element = chart();
    element.barDirection = 'col';
    element.grouping = 'clustered';

    const xml = serializeChartPart(element, 3);

    expect(xml).toContain('<c:barChart><c:barDir val="col"/>');
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:axId val="10000006"/>');
    expect(xml).toContain('<c:axId val="10000007"/>');
    expect(xml).toContain('Sheet1!$B$2:$B$3');
    expect(xml).toContain('<c:v>Revenue &lt;net&gt;</c:v>');
    expect(xml).toContain('<c:v>Q2 &amp; Q3</c:v>');
    expect(xml).toContain('<c:v>-20.5</c:v>');
    expect(xml).toContain('<a:srgbClr val="4F46E5"/>');
  });

  it('serializes line markers and grouping', () => {
    const element = chart('lineChart');
    element.grouping = 'stacked';
    element.marker = true;

    const xml = serializeChartPart(element, 1);

    expect(xml).toContain('<c:lineChart><c:grouping val="stacked"/>');
    expect(xml).toContain('<c:symbol val="circle"/>');
    expect(xml).toContain('<c:marker val="1"/>');
  });

  it.each([
    ['pieChart', ''],
    ['doughnutChart', '<c:holeSize val="65"/>'],
  ] as const)('serializes a native %s without axes', (chartType, hole) => {
    const element = chart(chartType);
    if (chartType === 'doughnutChart') element.holeSize = 65;

    const xml = serializeChartPart(element, 2);

    expect(xml).toContain(`<c:${chartType}>`);
    expect(xml).toContain(hole);
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).not.toContain('<c:valAx>');
  });

  it('serializes an accessible graphic frame with exact transform and relation', () => {
    const element = chart();
    const transform = element.authored.transform;
    if (transform === undefined) throw new Error('Expected chart transform');

    const xml = serializeChartFrame(element, transform, 4, 'rId7');

    expect(xml).toContain(
      '<p:cNvPr id="4" name="Revenue &amp; growth" descr="Chart &lt;description&gt;" title="Accessible &quot;chart&quot;" hidden="1"/>',
    );
    expect(xml).toContain(
      '<p:xfrm><a:off x="508000" y="762000"/><a:ext cx="6096000" cy="3048000"/></p:xfrm>',
    );
    expect(xml).toContain('r:id="rId7"');
  });
});
