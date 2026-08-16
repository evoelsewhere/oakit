import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

const STROKE_ONLY_PRESETS = [
  'arc',
  'leftBrace',
  'rightBrace',
  'bracePair',
  'leftBracket',
  'rightBracket',
  'bracketPair',
  'lineInv',
] as const;

interface ShapeXmlOptions {
  border?: boolean;
  custom?: 'filled' | 'stroke-only';
  fill?: boolean;
  id: number;
  placeholderType?: string;
  preset?: string;
  text?: string | null;
}

function shapeXml({
  border = false,
  custom,
  fill = false,
  id,
  placeholderType,
  preset,
  text,
}: ShapeXmlOptions): string {
  const geometry = custom
    ? `<a:custGeom>
        <a:avLst/><a:pathLst>
          <a:path w="100" h="100"${custom === 'stroke-only' ? ' fill="none"' : ''}>
            <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
            <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
          </a:path>
        </a:pathLst>
      </a:custGeom>`
    : `<a:prstGeom prst="${preset ?? 'rect'}"><a:avLst/></a:prstGeom>`;
  const textBody =
    text === undefined
      ? ''
      : `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${
          text === null ? '' : `<a:r><a:t>${text}</a:t></a:r>`
        }</a:p></p:txBody>`;

  return `<p:sp>
    <p:nvSpPr>
      <p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/>
      <p:nvPr>${placeholderType ? `<p:ph type="${placeholderType}"/>` : ''}</p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="${id * 1000}" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
      ${geometry}
      ${fill ? '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>' : ''}
      ${border ? '<a:ln w="12700"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:ln>' : ''}
    </p:spPr>
    ${textBody}
  </p:sp>`;
}

describe('PowerPoint shape classification through the public API', () => {
  it('distinguishes custom, stroke-only, object, and text-bearing shapes', async () => {
    const presetShapes = STROKE_ONLY_PRESETS.map((preset, index) =>
      shapeXml({ id: 610 + index, preset }),
    ).join('');
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${shapeXml({ custom: 'filled', id: 600, text: 'Custom text' })}
            ${shapeXml({ custom: 'stroke-only', id: 601, text: null })}
            ${presetShapes}
            ${shapeXml({ id: 620, preset: 'ellipse' })}
            ${shapeXml({ id: 621, placeholderType: 'obj' })}
            ${shapeXml({ id: 622 })}
            ${shapeXml({ id: 623, placeholderType: 'body' })}
            ${shapeXml({ fill: true, id: 624, placeholderType: 'body' })}
            ${shapeXml({ border: true, id: 625, placeholderType: 'body' })}
            ${shapeXml({ fill: true, id: 626, placeholderType: 'body', text: 'Filled text' })}
            ${shapeXml({ id: 627, placeholderType: 'body', preset: 'ellipse', text: 'Ellipse text' })}
          </p:spTree></p:cSld>
        </p:sld>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(byId['600']).toMatchObject({
      shapType: 'custom',
      type: 'shape',
    });
    expect(byId['600']?.type).toBe('shape');
    if (byId['600']?.type !== 'shape') {
      throw new Error('Expected the filled custom geometry to be a shape');
    }
    expect(byId['600'].content).toContain('Custom&nbsp;text');
    expect(byId['600']).not.toHaveProperty('strokeOnly');
    expect(byId['601']).toMatchObject({
      content: '',
      shapType: 'custom',
      strokeOnly: true,
      type: 'shape',
    });

    for (const [index, preset] of STROKE_ONLY_PRESETS.entries()) {
      expect(byId[String(610 + index)]).toMatchObject({
        content: '',
        shapType: preset,
        strokeOnly: true,
        type: 'shape',
      });
    }
    expect(byId['620']).toMatchObject({
      content: '',
      shapType: 'ellipse',
      type: 'shape',
    });
    expect(byId['620']).not.toHaveProperty('strokeOnly');
    expect(byId['621']).toMatchObject({ shapType: 'rect', type: 'shape' });
    expect(byId['622']).toMatchObject({ shapType: 'rect', type: 'shape' });
    expect(byId['623']).toMatchObject({ content: '', type: 'text' });
    expect(byId['624']).toMatchObject({
      content: '',
      fill: { type: 'color', value: '#123456' },
      shapType: 'rect',
      type: 'shape',
    });
    expect(byId['625']).toMatchObject({
      borderColor: '#654321',
      borderWidth: 1,
      content: '',
      shapType: 'rect',
      type: 'shape',
    });
    expect(byId['626']).toMatchObject({
      type: 'text',
    });
    expect(byId['626']?.type).toBe('text');
    if (byId['626']?.type !== 'text') {
      throw new Error('Expected the filled body placeholder to remain text');
    }
    expect(byId['626'].content).toContain('Filled&nbsp;text');
    expect(byId['627']).toMatchObject({
      shapType: 'ellipse',
      type: 'shape',
    });
    expect(byId['627']?.type).toBe('shape');
    if (byId['627']?.type !== 'shape') {
      throw new Error('Expected the ellipse body placeholder to be a shape');
    }
    expect(byId['627'].content).toContain('Ellipse&nbsp;text');
  });
});
