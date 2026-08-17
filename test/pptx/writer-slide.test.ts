import { describe, expect, it } from 'vitest';

import type {
  PptxSceneElement,
  PptxSceneSlide,
  PptxSceneTextElement,
} from '../../src/formats/pptx/scene-types';
import { createFieldIdAllocator } from '../../src/formats/pptx/writer/identifiers';
import { serializeSlide } from '../../src/formats/pptx/writer/slide';

const PREFIX =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const ROOT =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function textElement(key: string, text = key): PptxSceneTextElement {
  return {
    authored: {
      transform: { height: 40, width: 100, x: 10, y: 20 },
    },
    key,
    resolved: { hidden: false },
    text: {
      body: {},
      paragraphs: [
        {
          children: [{ key: `${key}-run`, text, type: 'run' }],
          key: `${key}-paragraph`,
        },
      ],
    },
    type: 'text',
  };
}

function slide(elements: PptxSceneElement[] = []): PptxSceneSlide {
  return { elements, key: 'slide-1' };
}

describe('PowerPoint slide serialization', () => {
  it('serializes the literal minimum slide shape tree', () => {
    expect(serializeSlide(slide(), createFieldIdAllocator())).toBe(
      `${PREFIX}><p:cSld><p:spTree>${ROOT}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );
  });

  it('maps explicit hidden and visible slide state to show', () => {
    const hidden = serializeSlide(
      { ...slide(), hidden: true },
      createFieldIdAllocator(),
    );
    const visible = serializeSlide(
      { ...slide(), hidden: false },
      createFieldIdAllocator(),
    );
    const unspecified = serializeSlide(slide(), createFieldIdAllocator());

    expect(hidden.slice(0, `${PREFIX} show="0">`.length)).toBe(
      `${PREFIX} show="0">`,
    );
    expect(visible.slice(0, `${PREFIX} show="1">`.length)).toBe(
      `${PREFIX} show="1">`,
    );
    expect(unspecified.slice(0, `${PREFIX}>`.length)).toBe(`${PREFIX}>`);
  });

  it('escapes the authored common-slide name', () => {
    expect(
      serializeSlide(
        { ...slide(), name: `Slide <&"' _x0041_` },
        createFieldIdAllocator(),
      ),
    ).toContain('<p:cSld name="Slide &lt;&amp;&quot;&apos; _x005F_x0041_">');
  });

  it('serializes a solid slide background before the shape tree', () => {
    expect(
      serializeSlide(
        { ...slide(), backgroundColor: '#0f172a' },
        createFieldIdAllocator(),
      ),
    ).toContain(
      '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>',
    );
  });

  it('assigns owner-scoped shape IDs in authored z-order', () => {
    const xml = serializeSlide(
      slide([textElement('first', 'First'), textElement('second', 'Second')]),
      createFieldIdAllocator(),
    );

    expect(xml).toContain('<p:cNvPr id="2" name="Text Box 2"/>');
    expect(xml).toContain('<p:cNvPr id="3" name="Text Box 3"/>');
    expect(xml).toContain('</p:sp><p:sp>');
    expect(xml.indexOf('<a:t>First</a:t>')).toBeLessThan(
      xml.indexOf('<a:t>Second</a:t>'),
    );
  });

  it('shares field allocation across shapes and slides in one write', () => {
    const first = textElement('first');
    first.text.paragraphs[0] = {
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
    const second = textElement('second');
    second.text.paragraphs[0] = {
      children: [
        {
          fieldType: 'datetime',
          key: 'field-2',
          text: 'Now',
          type: 'field',
        },
      ],
      key: 'paragraph-2',
    };
    const allocation = createFieldIdAllocator();

    const firstXml = serializeSlide(slide([first]), allocation);
    const secondXml = serializeSlide(
      { elements: [second], key: 'slide-2' },
      allocation,
    );
    expect(firstXml).toContain('id="{00000000-0000-0000-0000-000000000001}"');
    expect(secondXml).toContain('id="{00000000-0000-0000-0000-000000000002}"');
    expect(secondXml).toContain('<p:cNvPr id="2" name="Text Box 2"/>');
  });

  it('rejects preservation-only elements at the serializer boundary', () => {
    const unsupported: PptxSceneElement = {
      authored: {},
      feature: 'chart',
      key: 'chart-1',
      resolved: { hidden: false },
      type: 'unsupported',
    };

    expect(() =>
      serializeSlide(slide([unsupported]), createFieldIdAllocator()),
    ).toThrow(
      new TypeError('PowerPoint slide writer accepts text elements only'),
    );
  });

  it('rejects a text element missing authored geometry', () => {
    const element = textElement('text-1');
    element.authored = {};

    expect(() =>
      serializeSlide(slide([element]), createFieldIdAllocator()),
    ).toThrow(
      new TypeError(
        'PowerPoint slide writer requires an authored text transform',
      ),
    );
  });
});
