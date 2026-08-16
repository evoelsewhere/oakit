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

const RICH_CONTENT_SLIDE = `
  <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Adjusted shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
          <a:solidFill><a:srgbClr val="336699"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="3" name="Revenue table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1"/>
            <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
            <a:tr h="457200">
              <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"><a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill></a:tcPr></a:tc>
              <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>
            </a:tr>
            <a:tr h="457200">
              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Q1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>42</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="4" name="Revenue chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="914400"/><a:ext cx="2743200" cy="1828800"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rIdChart"/>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree></p:cSld>
    <p:transition spd="fast" advClick="0" advTm="2500"><p:wipe dir="l"/></p:transition>
  </p:sld>`;

const BAR_CHART = `
  <c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}">
    <c:chart><c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser>
        <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:strCache>
          <c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache>
          <c:pt idx="0"><c:v>42</c:v></c:pt><c:pt idx="1"><c:v>55</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
        <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
      </c:ser>
    </c:barChart></c:plotArea></c:chart>
  </c:chartSpace>`;

describe('PPTX rich content through the public API', () => {
  it('parses adjusted shapes, tables, charts, and transitions together', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
        <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:defaultTextStyle>
      </p:presentation>`,
      'ppt/slides/slide1.xml': RICH_CONTENT_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdChart" Type="${OFFICE_REL_TYPE}chart" Target="../charts/chart1.xml"/>
        </Relationships>`,
      'ppt/charts/chart1.xml': BAR_CHART,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const slide = result.slides[0];
    const shape = slide?.elements.find((element) => element.type === 'shape');
    const table = slide?.elements.find((element) => element.type === 'table');
    const chart = slide?.elements.find((element) => element.type === 'chart');

    expect(shape).toMatchObject({
      id: '2',
      shapType: 'roundRect',
      rotate: 90,
      keypoints: { adj: 0.5 },
    });
    expect(table?.type).toBe('table');
    if (table?.type !== 'table') throw new Error('Expected a table');
    expect(table.data).toHaveLength(2);
    expect(table.data[0]?.[0]).toMatchObject({
      colSpan: 2,
      fillColor: '#D9EAF7',
      vAlign: 'mid',
    });
    expect(table.data[0]?.[0]?.text).toContain('Quarterly&nbsp;revenue');
    expect(table.data[0]?.[0]?.text).toContain('font-size: 24pt');
    expect(table.data[0]?.[1]).toMatchObject({ hMerge: 1 });
    expect(table.data[1]?.[1]?.text).toContain('42');
    expect(table.colWidths).toEqual([72, 72]);
    expect(chart?.type).toBe('chart');
    if (chart?.type !== 'chart') throw new Error('Expected a chart');
    expect(chart).toMatchObject({
      id: '4',
      chartType: 'barChart',
      barDir: 'col',
      grouping: 'clustered',
      colors: ['#4472C4'],
    });
    expect(chart.data).toEqual([
      {
        key: 'Revenue',
        values: [
          { x: '0', y: 42 },
          { x: '1', y: 55 },
        ],
        xlabels: { 0: 'Q1', 1: 'Q2' },
      },
    ]);
    expect(slide?.transition).toMatchObject({
      type: 'wipe',
      duration: 500,
      direction: 'l',
      autoNextAfter: 2500,
    });
  });
});
