import { describe, expect, it, vi } from 'vitest';

import type {
  PptxSceneTextElement,
  PptxSceneTransform,
} from '../../src/formats/pptx/scene-types';
import type { PptxTextSerializationContext } from '../../src/formats/pptx/writer/text-node';
import {
  serializeShapeTransform,
  serializeTextShape,
} from '../../src/formats/pptx/writer/text-shape';

function context(fieldId = '{00000000-0000-0000-0000-000000000001}') {
  return {
    allocateFieldId: vi.fn(() => fieldId),
  } satisfies PptxTextSerializationContext;
}

function textElement(): PptxSceneTextElement {
  return {
    authored: {},
    key: 'text-1',
    resolved: { hidden: false },
    text: {
      body: {},
      paragraphs: [
        {
          children: [{ key: 'run-1', text: 'Hello', type: 'run' }],
          key: 'paragraph-1',
        },
      ],
    },
    type: 'text',
  };
}

const TRANSFORM: PptxSceneTransform = {
  height: 40,
  width: 100,
  x: 10,
  y: 20,
};

describe('PowerPoint text-shape serialization', () => {
  it('converts a basic authored transform to exact EMUs', () => {
    expect(serializeShapeTransform(TRANSFORM)).toBe(
      '<a:xfrm><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm>',
    );
  });

  it('serializes rotation and explicit flip values in deterministic order', () => {
    expect(
      serializeShapeTransform({
        ...TRANSFORM,
        flipHorizontal: false,
        flipVertical: true,
        rotation: -45.5,
      }),
    ).toBe(
      '<a:xfrm rot="-2730000" flipH="0" flipV="1"><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm>',
    );
  });

  it('supports negative positions and rounded fractional sizes', () => {
    expect(
      serializeShapeTransform({
        height: 0.00004,
        width: 0.00004,
        x: -0.5,
        y: -1,
      }),
    ).toBe(
      '<a:xfrm><a:off x="-6350" y="-12700"/><a:ext cx="1" cy="1"/></a:xfrm>',
    );
  });

  it('uses a deterministic required name when authored name is absent', () => {
    expect(serializeTextShape(textElement(), TRANSFORM, 2, context())).toBe(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Text Box 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>',
    );
  });

  it('escapes authored nonvisual metadata and Office sequences', () => {
    const element = textElement();
    element.name = `Name <&"' _x0041_`;
    element.description = `Description <&"'`;
    element.title = `Title <&"'`;
    element.authored.hidden = true;

    expect(serializeTextShape(element, TRANSFORM, 7, context())).toContain(
      '<p:cNvPr id="7" name="Name &lt;&amp;&quot;&apos; _x005F_x0041_" descr="Description &lt;&amp;&quot;&apos;" title="Title &lt;&amp;&quot;&apos;" hidden="1"/>',
    );
  });

  it('serializes rich fill, border, and geometry styling', () => {
    const element = textElement();
    element.authored.fillColor = '#0f172a';
    element.authored.geometry = 'roundRect';
    element.authored.lineColor = '#38bdf8';
    element.authored.lineWidth = 1.5;

    expect(serializeTextShape(element, TRANSFORM, 2, context())).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill></a:ln>',
    );
  });

  it('uses a deterministic black border when only width is authored', () => {
    const element = textElement();
    element.authored.lineWidth = 2;

    expect(serializeTextShape(element, TRANSFORM, 2, context())).toContain(
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
    );
  });

  it('uses the OOXML default width when only border color is authored', () => {
    const element = textElement();
    element.authored.lineColor = '#F97316';

    expect(serializeTextShape(element, TRANSFORM, 2, context())).toContain(
      '<a:ln><a:solidFill><a:srgbClr val="F97316"/></a:solidFill></a:ln>',
    );
  });

  it('distinguishes explicit visible state from authored absence', () => {
    const visible = textElement();
    visible.authored.hidden = false;
    expect(serializeTextShape(visible, TRANSFORM, 2, context())).toContain(
      ' hidden="0"/>',
    );

    const inheritedPreview = textElement();
    inheritedPreview.resolved.hidden = true;
    expect(
      serializeTextShape(inheritedPreview, TRANSFORM, 2, context()),
    ).not.toContain(' hidden=');
  });

  it('uses the supplied authored transform instead of the resolved preview', () => {
    const element = textElement();
    element.resolved.transform = {
      height: 400,
      width: 500,
      x: 600,
      y: 700,
    };

    const xml = serializeTextShape(element, TRANSFORM, 2, context());
    expect(xml).toContain(
      '<a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/>',
    );
    expect(xml).not.toContain('7620000');
  });

  it('passes one local field allocation context through the text body', () => {
    const element = textElement();
    element.text.paragraphs[0] = {
      children: [
        {
          fieldType: 'slidenum',
          key: 'field-1',
          text: '1',
          type: 'field',
        },
      ],
      key: 'paragraph-1',
    };
    const allocation = context('{field-local}');

    expect(serializeTextShape(element, TRANSFORM, 3, allocation)).toContain(
      '<a:fld id="{field-local}" type="slidenum">',
    );
    expect(allocation.allocateFieldId).toHaveBeenCalledTimes(1);
  });
});
