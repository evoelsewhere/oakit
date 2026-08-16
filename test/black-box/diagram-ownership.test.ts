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

const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DIAGRAM_DRAWING_NS =
  'http://schemas.microsoft.com/office/drawing/2008/diagram';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

function diagramFrame(
  id: string,
  relationshipId: string | undefined,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  const dataRelationship = relationshipId ? ` r:dm="${relationshipId}"` : '';
  return `
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Diagram ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm>
      <a:graphic><a:graphicData uri="${DIAGRAM_URI}"><dgm:relIds${dataRelationship}/></a:graphicData></a:graphic>
    </p:graphicFrame>`;
}

function shapeTree(frames: string): string {
  return `<p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${frames}</p:spTree></p:cSld>`;
}

function diagramData(label: string): string {
  return `
    <dgm:dataModel xmlns:dgm="${DIAGRAM_NS}" xmlns:a="${DRAWING_NS}" xmlns:dsp="${DIAGRAM_DRAWING_NS}">
      <dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>${label}</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>
      <dgm:extLst><a:ext uri="drawing"><dsp:dataModelExt relId="rIdDrawing"/></a:ext></dgm:extLst>
    </dgm:dataModel>`;
}

function diagramDrawing(
  shapeId: string,
  imageRelationshipId: string,
  textId?: string,
): string {
  const textShape = textId
    ? `<dsp:sp>
        <dsp:nvSpPr><dsp:cNvPr id="${textId}" name="Diagram text"/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr>
        <dsp:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></dsp:spPr>
        <dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Rendered text</a:t></a:r></a:p></dsp:txBody>
      </dsp:sp>`
    : '';
  return `
    <dsp:drawing xmlns:dsp="${DIAGRAM_DRAWING_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
      <dsp:spTree>
        <dsp:sp>
          <dsp:nvSpPr><dsp:cNvPr id="${shapeId}" name="Diagram image"/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr>
          <dsp:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:blipFill><a:blip r:embed="${imageRelationshipId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
          </dsp:spPr>
        </dsp:sp>
        ${textShape}
      </dsp:spTree>
    </dsp:drawing>`;
}

function diagramRelationships(drawing: string): string {
  return `
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdDrawing" Type="${OFFICE_REL_TYPE}diagramDrawing" Target="${drawing}"/>
    </Relationships>`;
}

function drawingRelationships(image: string): string {
  return `
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/${image}"/>
    </Relationships>`;
}

describe('PowerPoint diagram ownership through the public API', () => {
  it('resolves data, drawings, media, text, and empty diagrams from their owners', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(
            diagramFrame(
              '30',
              'rIdOwnedData',
              914400,
              1828800,
              2743200,
              1371600,
            ) + diagramFrame('33', undefined, 0, 0, 457200, 457200),
          )}
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdOwnedData" Type="${OFFICE_REL_TYPE}diagramData" Target="../diagrams/slide-data.xml"/>
        </Relationships>`,
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(diagramFrame('31', 'rIdOwnedData', 0, 457200, 1828800, 914400))}
        </p:sldLayout>`,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdMaster" Type="${OFFICE_REL_TYPE}slideMaster" Target="../slideMasters/slideMaster1.xml"/>
          <Relationship Id="rIdOwnedData" Type="${OFFICE_REL_TYPE}diagramData" Target="../diagrams/layout-data.xml"/>
        </Relationships>`,
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${OFFICE_REL_NS}">
          ${shapeTree(diagramFrame('32', 'rIdOwnedData', 457200, 0, 914400, 457200))}
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="../theme/theme1.xml"/>
          <Relationship Id="rIdOwnedData" Type="${OFFICE_REL_TYPE}diagramData" Target="../diagrams/master-data.xml"/>
        </Relationships>`,
      'ppt/diagrams/slide-data.xml': diagramData('Slide data'),
      'ppt/diagrams/layout-data.xml': diagramData('Layout data'),
      'ppt/diagrams/master-data.xml': diagramData('Master data'),
      'ppt/diagrams/_rels/slide-data.xml.rels':
        diagramRelationships('slide-drawing.xml'),
      'ppt/diagrams/_rels/layout-data.xml.rels':
        diagramRelationships('layout-drawing.xml'),
      'ppt/diagrams/_rels/master-data.xml.rels':
        diagramRelationships('master-drawing.xml'),
      'ppt/diagrams/slide-drawing.xml': diagramDrawing('40', 'rIdImage', '41'),
      'ppt/diagrams/layout-drawing.xml': diagramDrawing('50', 'rIdImage'),
      'ppt/diagrams/master-drawing.xml': diagramDrawing('60', 'rIdImage'),
      'ppt/diagrams/_rels/slide-drawing.xml.rels':
        drawingRelationships('slide.png'),
      'ppt/diagrams/_rels/layout-drawing.xml.rels':
        drawingRelationships('layout.png'),
      'ppt/diagrams/_rels/master-drawing.xml.rels':
        drawingRelationships('master.png'),
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const slideDiagrams = result.slides[0]?.elements ?? [];
    const layoutDiagrams = result.slides[0]?.layoutElements ?? [];
    const slideDiagram = slideDiagrams.find((element) => element.id === '30');
    const emptyDiagram = slideDiagrams.find((element) => element.id === '33');
    const layoutDiagram = layoutDiagrams.find((element) => element.id === '31');
    const masterDiagram = layoutDiagrams.find((element) => element.id === '32');
    if (slideDiagram?.type !== 'diagram')
      throw new Error('Expected slide diagram');
    if (emptyDiagram?.type !== 'diagram')
      throw new Error('Expected empty diagram');
    if (layoutDiagram?.type !== 'diagram')
      throw new Error('Expected layout diagram');
    if (masterDiagram?.type !== 'diagram')
      throw new Error('Expected master diagram');

    const {
      elements: slideElements,
      order: slideOrder,
      ...slideData
    } = slideDiagram;
    expect(slideData).toEqual({
      type: 'diagram',
      id: '30',
      left: 72,
      top: 144,
      width: 216,
      height: 108,
      textList: ['Slide data'],
    });
    expect(slideElements).toHaveLength(2);
    expect(slideElements[0]).toMatchObject({
      type: 'shape',
      id: '40',
      fill: {
        type: 'image',
        value: { ref: 'ppt/media/slide.png' },
      },
    });
    expect(slideElements[1]).toMatchObject({ type: 'text', id: '41' });
    if (slideElements[1]?.type !== 'text') {
      throw new Error('Expected rendered diagram text');
    }
    expect(slideElements[1].content).toContain('Rendered&nbsp;text');
    expect(slideOrder).toBeGreaterThan(0);

    const { order: emptyOrder, ...emptyData } = emptyDiagram;
    expect(emptyData).toEqual({
      type: 'diagram',
      id: '33',
      left: 0,
      top: 0,
      width: 36,
      height: 36,
      elements: [],
      textList: [],
    });
    expect(emptyOrder).toBeGreaterThan(0);

    expectOwnedDiagram(layoutDiagram, {
      id: '31',
      left: 0,
      top: 36,
      width: 144,
      height: 72,
      text: 'Layout data',
      shapeId: '50',
      image: 'layout.png',
    });
    expectOwnedDiagram(masterDiagram, {
      id: '32',
      left: 36,
      top: 0,
      width: 72,
      height: 36,
      text: 'Master data',
      shapeId: '60',
      image: 'master.png',
    });
  });
});

interface OwnedDiagramExpectation {
  height: number;
  id: string;
  image: string;
  left: number;
  shapeId: string;
  text: string;
  top: number;
  width: number;
}

function expectOwnedDiagram(
  diagram: Extract<
    NonNullable<
      Awaited<ReturnType<typeof parsePptx>>['slides'][number]
    >['elements'][number],
    { type: 'diagram' }
  >,
  expected: OwnedDiagramExpectation,
): void {
  const { elements, order, ...data } = diagram;
  expect(data).toEqual({
    type: 'diagram',
    id: expected.id,
    left: expected.left,
    top: expected.top,
    width: expected.width,
    height: expected.height,
    textList: [expected.text],
  });
  expect(elements).toHaveLength(1);
  expect(elements[0]).toMatchObject({
    type: 'shape',
    id: expected.shapeId,
    fill: {
      type: 'image',
      value: { ref: `ppt/media/${expected.image}` },
    },
  });
  expect(order).toBeGreaterThan(0);
}
