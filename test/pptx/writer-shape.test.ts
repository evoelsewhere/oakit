import { describe, expect, it } from 'vitest';

import type {
  PptxSceneShapeElement,
  PptxSceneTransform,
} from '../../src/formats/pptx/scene-types';
import {
  serializeNativeShapeProperties,
  serializeShape,
  serializeShapeNonVisualProperties,
  serializeShapeTransform,
} from '../../src/formats/pptx/writer/shape';

const TRANSFORM: PptxSceneTransform = {
  height: 40,
  width: 100,
  x: 10,
  y: 20,
};

function shapeElement(): PptxSceneShapeElement {
  return {
    authored: {},
    key: 'shape-1',
    resolved: { hidden: false },
    type: 'shape',
  };
}

describe('native PowerPoint shape serialization', () => {
  it('serializes a complete native shape without a text body', () => {
    expect(serializeShape(shapeElement(), TRANSFORM, 2)).toBe(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:sp>',
    );
  });

  it('serializes signed transforms, rotation, and flips', () => {
    expect(
      serializeShapeTransform({
        ...TRANSFORM,
        flipHorizontal: false,
        flipVertical: true,
        rotation: -12.5,
        x: -10,
      }),
    ).toBe(
      '<a:xfrm rot="-750000" flipH="0" flipV="1"><a:off x="-127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm>',
    );
  });

  it('escapes metadata and distinguishes native shapes from text boxes', () => {
    const element = shapeElement();
    element.name = `Shape <&"' _x0041_`;
    element.description = `Description <&"'`;
    element.title = `Title <&"'`;
    element.authored.hidden = false;

    expect(serializeShapeNonVisualProperties(element, 7, false)).toBe(
      '<p:nvSpPr><p:cNvPr id="7" name="Shape &lt;&amp;&quot;&apos; _x005F_x0041_" descr="Description &lt;&amp;&quot;&apos;" title="Title &lt;&amp;&quot;&apos;" hidden="0"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>',
    );
    expect(serializeShapeNonVisualProperties(element, 7, true)).toContain(
      '<p:cNvSpPr txBox="1"/>',
    );
  });

  it('serializes geometry, fill, line color, and line width', () => {
    const element = shapeElement();
    element.authored.fillColor = '#0f172a';
    element.authored.geometry = 'ellipse';
    element.authored.lineColor = '#38bdf8';
    element.authored.lineWidth = 1.5;

    expect(serializeNativeShapeProperties(element, TRANSFORM)).toContain(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill></a:ln>',
    );
  });

  it('uses a black line when only width is authored', () => {
    const element = shapeElement();
    element.authored.lineWidth = 2;

    expect(serializeNativeShapeProperties(element, TRANSFORM)).toContain(
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
    );
  });

  it('uses the OOXML default width when only line color is authored', () => {
    const element = shapeElement();
    element.authored.lineColor = '#F97316';

    expect(serializeNativeShapeProperties(element, TRANSFORM)).toContain(
      '<a:ln><a:solidFill><a:srgbClr val="F97316"/></a:solidFill></a:ln>',
    );
  });
});
