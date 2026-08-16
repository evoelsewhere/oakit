import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function chartFrame(
  id: string,
  relationshipId: string | undefined,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  const relationship = relationshipId ? ` r:id="${relationshipId}"` : '';
  return `
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm>
      <a:graphic><a:graphicData uri="${CHART_URI}"><c:chart${relationship}/></a:graphicData></a:graphic>
    </p:graphicFrame>`;
}

function chartPart(
  type: string,
  value: number,
  options = '',
  series = `
    <c:ser>
      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${type}</c:v></c:pt></c:strCache></c:strRef></c:tx>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Only</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>${value}</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>`,
): string {
  return `
    <c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}">
      <c:chart><c:plotArea><c:${type}>${options}${series}</c:${type}></c:plotArea></c:chart>
    </c:chartSpace>`;
}

function shapeTree(frames: string): string {
  return `<p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${frames}</p:spTree></p:cSld>`;
}

describe('PowerPoint chart ownership through the public API', () => {
  it('resolves colliding relationship ids from the part that owns each chart', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(chartFrame('30', 'rIdOwnedChart', 914400, 1828800, 2743200, 1371600))}
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdOwnedChart" Type="${OFFICE_REL_TYPE}chart" Target="../charts/slide.xml"/>
        </Relationships>`,
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(chartFrame('31', 'rIdOwnedChart', 0, 457200, 1828800, 914400))}
        </p:sldLayout>`,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdMaster" Type="${OFFICE_REL_TYPE}slideMaster" Target="../slideMasters/slideMaster1.xml"/>
          <Relationship Id="rIdOwnedChart" Type="${OFFICE_REL_TYPE}chart" Target="../charts/layout.xml"/>
        </Relationships>`,
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(chartFrame('32', 'rIdOwnedChart', 457200, 0, 914400, 457200))}
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="../theme/theme1.xml"/>
          <Relationship Id="rIdOwnedChart" Type="${OFFICE_REL_TYPE}chart" Target="../charts/master.xml"/>
        </Relationships>`,
      'ppt/charts/slide.xml': chartPart(
        'barChart',
        11,
        '<c:barDir val="bar"/><c:grouping val="clustered"/>',
      ),
      'ppt/charts/layout.xml': chartPart(
        'lineChart',
        22,
        '<c:grouping val="stacked"/><c:marker/>',
      ),
      'ppt/charts/master.xml': chartPart(
        'doughnutChart',
        33,
        '<c:holeSize val="65"/>',
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const slideChart = result.slides[0]?.elements[0];
    const [layoutChart, masterChart] = result.slides[0]?.layoutElements ?? [];
    if (slideChart?.type !== 'chart') throw new Error('Expected slide chart');
    if (layoutChart?.type !== 'chart') throw new Error('Expected layout chart');
    if (masterChart?.type !== 'chart') throw new Error('Expected master chart');
    const { order: slideOrder, ...slideChartData } = slideChart;
    const { order: layoutOrder, ...layoutChartData } = layoutChart;
    const { order: masterOrder, ...masterChartData } = masterChart;

    expect(slideChartData).toEqual({
      type: 'chart',
      id: '30',
      left: 72,
      top: 144,
      width: 216,
      height: 108,
      data: [
        {
          key: 'barChart',
          values: [{ x: '0', y: 11 }],
          xlabels: { 0: 'Only' },
        },
      ],
      colors: [''],
      chartType: 'barChart',
      barDir: 'bar',
      grouping: 'clustered',
    });
    expect(layoutChartData).toEqual({
      type: 'chart',
      id: '31',
      left: 0,
      top: 36,
      width: 144,
      height: 72,
      data: [
        {
          key: 'lineChart',
          values: [{ x: '0', y: 22 }],
          xlabels: { 0: 'Only' },
        },
      ],
      colors: [''],
      chartType: 'lineChart',
      grouping: 'stacked',
      marker: true,
    });
    expect(masterChartData).toEqual({
      type: 'chart',
      id: '32',
      left: 36,
      top: 0,
      width: 72,
      height: 36,
      data: [
        {
          key: 'doughnutChart',
          values: [{ x: '0', y: 33 }],
          xlabels: { 0: 'Only' },
        },
      ],
      colors: [],
      chartType: 'doughnutChart',
      holeSize: '65',
    });
    expect(slideOrder).toBeGreaterThan(0);
    expect(layoutOrder).toBeGreaterThan(0);
    expect(masterOrder).toBeGreaterThan(0);
  });

  it('omits unresolved and unsupported charts without hiding a valid scatter chart', async () => {
    const invalidFrames = [
      chartFrame('40', undefined, 0, 0, 914400, 914400),
      chartFrame('41', 'rIdUnknown', 914400, 0, 914400, 914400),
      chartFrame('42', 'rIdEmpty', 1828800, 0, 914400, 914400),
      chartFrame('43', 'rIdNoPlot', 2743200, 0, 914400, 914400),
      chartFrame('44', 'rIdUnsupported', 3657600, 0, 914400, 914400),
    ].join('');
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(`${invalidFrames}${chartFrame('45', 'rIdScatter', 4572000, 914400, 1828800, 1371600)}`)}
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdEmpty" Type="${OFFICE_REL_TYPE}chart" Target="../charts/empty.xml"/>
          <Relationship Id="rIdNoPlot" Type="${OFFICE_REL_TYPE}chart" Target="../charts/no-plot.xml"/>
          <Relationship Id="rIdUnsupported" Type="${OFFICE_REL_TYPE}chart" Target="../charts/unsupported.xml"/>
          <Relationship Id="rIdScatter" Type="${OFFICE_REL_TYPE}chart" Target="../charts/scatter.xml"/>
        </Relationships>`,
      'ppt/charts/empty.xml': '<root/>',
      'ppt/charts/no-plot.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart/></c:chartSpace>`,
      'ppt/charts/unsupported.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:unknownChart><c:ser/></c:unknownChart></c:plotArea></c:chart></c:chartSpace>`,
      'ppt/charts/scatter.xml': chartPart(
        'scatterChart',
        0,
        '<c:scatterStyle val="smoothMarker"/>',
        `<c:ser>
          <c:xVal><c:numRef><c:numCache><c:pt idx="0"><c:v>1.5</c:v></c:pt></c:numCache></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:numCache><c:pt idx="0"><c:v>2.5</c:v></c:pt></c:numCache></c:numRef></c:yVal>
        </c:ser>`,
      ),
    });

    const elements = resultElements(
      await parsePptx(input, { errorMode: 'strict' }),
    );
    const scatterChart = elements[0];
    if (scatterChart?.type !== 'chart') {
      throw new Error('Expected one scatter chart');
    }
    const { order, ...scatterChartData } = scatterChart;

    expect(elements).toHaveLength(1);
    expect(scatterChartData).toEqual({
      type: 'chart',
      id: '45',
      left: 360,
      top: 72,
      width: 144,
      height: 108,
      data: [[1.5], [2.5]],
      colors: [''],
      chartType: 'scatterChart',
      style: 'smoothMarker',
    });
    expect(order).toBeGreaterThan(0);
  });
});

function resultElements(
  result: Awaited<ReturnType<typeof parsePptx>>,
): NonNullable<(typeof result.slides)[number]>['elements'] {
  return result.slides[0]?.elements ?? [];
}
