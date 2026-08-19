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

    expect(xml).toContain(
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );
    expect(xml).toContain('<c:barChart><c:barDir val="col"/>');
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:axId val="10000006"/>');
    expect(xml).toContain('<c:axId val="10000007"/>');
    expect(xml).toContain('Sheet1!$B$2:$B$3');
    expect(xml).toContain('<c:v>Revenue &lt;net&gt;</c:v>');
    expect(xml).toContain('<c:v>Q2 &amp; Q3</c:v>');
    expect(xml).toContain('<c:v>-20.5</c:v>');
    expect(xml.match(/<a:srgbClr val="4F46E5"\/>/g)).toHaveLength(2);
    expect(xml).not.toContain('4f46e5');
    expect(xml).not.toContain('val="#4F46E5"');
    expect(xml).toContain('<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>');
    expect(xml).toContain(
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache>',
    );
    expect(xml).toContain(
      '<c:catAx><c:axId val="10000006"/><c:scaling><c:orientation val="minMax"/></c:scaling>',
    );
    expect(xml).toContain(
      '<c:valAx><c:axId val="10000007"/><c:scaling><c:orientation val="minMax"/></c:scaling>',
    );
    expect(xml).not.toContain('<c:marker>');
    expect(xml).not.toContain('Stryker was here!');
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

  it('serializes explicit false and default chart options', () => {
    const line = chart('lineChart');
    line.marker = false;
    const lineXml = serializeChartPart(line, 1);
    expect(lineXml).toContain('<c:grouping val="standard"/>');
    expect(lineXml).toContain('<c:symbol val="none"/>');
    expect(lineXml).toContain('<c:marker val="0"/>');
    const defaultLine = chart('lineChart');
    const defaultLineXml = serializeChartPart(defaultLine, 1);
    expect(defaultLineXml).toContain('<c:symbol val="none"/>');
    expect(defaultLineXml).toContain('<c:marker val="0"/>');

    const bar = chart('barChart');
    const defaultBarXml = serializeChartPart(bar, 1);
    expect(defaultBarXml).toContain('<c:barDir val="col"/>');
    expect(defaultBarXml).toContain('<c:grouping val="clustered"/>');
    bar.barDirection = 'bar';
    bar.grouping = 'stacked';
    const explicitBarXml = serializeChartPart(bar, 1);
    expect(explicitBarXml).toContain('<c:barDir val="bar"/>');
    expect(explicitBarXml).toContain('<c:grouping val="stacked"/>');
  });

  it('omits series styling when a color is not authored', () => {
    const element = chart();
    delete element.series[0]!.color;
    element.series.push({
      categories: ['Q1', 'Q2 & Q3'],
      key: 'series-2',
      name: 'Second',
      values: [5, 6],
    });

    const xml = serializeChartPart(element, 1);

    expect(xml).not.toContain('<c:spPr>');
    expect(xml).not.toContain('Stryker was here!');
    expect(xml.match(/<c:ser>/g)).toHaveLength(2);
    expect(xml).toContain('</c:ser><c:ser>');
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
    if (chartType === 'pieChart') expect(xml).not.toContain('<c:holeSize');
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).not.toContain('<c:valAx>');
    expect(xml).not.toContain('Stryker was here!');
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
    expect(xml).toContain(
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>',
    );
  });

  it('uses the deterministic chart name and explicit visible state defaults', () => {
    const element = chart();
    delete element.name;
    element.authored.hidden = false;
    const transform = element.authored.transform;
    if (transform === undefined) throw new Error('Expected chart transform');

    const xml = serializeChartFrame(element, transform, 9, 'rId2');

    expect(xml).toContain('<p:cNvPr id="9" name="Chart 9"');
    expect(xml).toContain('hidden="0"');
  });
});
