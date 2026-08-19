import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

function shape(
  id: string,
  x: number,
  y: number,
  width: number,
  height: number,
  rotation = 0,
): string {
  return `
    <p:sp>
      <p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm rot="${rotation}"><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </p:spPr>
    </p:sp>`;
}

function group(
  id: string,
  transform: string,
  children: string,
  attributes = '',
): string {
  return `
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm ${attributes}>${transform}</a:xfrm></p:grpSpPr>
      ${children}
    </p:grpSp>`;
}

function findGroup(
  elements: Awaited<ReturnType<typeof parsePptx>>['slides'][number]['elements'],
  id: string,
) {
  const element = elements.find((candidate) => candidate.id === id);
  expect(element?.type).toBe('group');
  if (element?.type !== 'group') throw new Error(`Expected group ${id}`);
  return element;
}

describe('PowerPoint group transforms through the public API', () => {
  it('applies non-uniform, uniform, zero, and nested coordinate systems exactly', async () => {
    const deeplyNested = group(
      '212',
      '<a:off x="228600" y="228600"/><a:ext cx="457200" cy="457200"/><a:chOff x="0" y="0"/><a:chExt cx="457200" cy="457200"/>',
      shape('213', 228600, 114300, 114300, 228600),
    );
    const nested = group(
      '210',
      '<a:off x="914400" y="457200"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/>',
      [shape('211', 0, 0, 457200, 457200), deeplyNested].join(''),
    );
    const nonUniform = group(
      '200',
      '<a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="2743200"/><a:chOff x="457200" y="228600"/><a:chExt cx="1828800" cy="914400"/>',
      [
        shape('201', 914400, 457200, 457200, 228600),
        nested,
        shape('202', 1371600, 685800, 914400, 457200, 5400000),
        shape('203', 0, 0, 457200, 914400, 16200000),
      ].join(''),
      'flipV="1" flipH="1" rot="5400000"',
    );
    const uniform = group(
      '220',
      '<a:off x="0" y="3657600"/><a:ext cx="1828800" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/>',
      shape('221', 0, 0, 457200, 228600, 5400000),
    );
    const zeroScale = group(
      '230',
      '<a:off x="2743200" y="3657600"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>',
      shape('231', 914400, 914400, 457200, 457200),
    );
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${nonUniform}
            ${uniform}
            ${zeroScale}
            <p:grpSp><p:nvGrpSpPr><p:cNvPr id="240" name="Missing transform"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp>
          </p:spTree></p:cSld>
        </p:sld>`,
    });

    const document = await parsePptx(input, { errorMode: 'strict' });
    const elements = document.slides[0]?.elements ?? [];
    const outer = findGroup(elements, '200');

    expect(outer).toMatchObject({
      childSpace: { height: 72, width: 144, x: 36, y: 18 },
      isFlipH: true,
      isFlipV: true,
      height: 216,
      left: 72,
      rotate: 90,
      top: 144,
      width: 288,
    });
    expect(outer.elements.map((element) => element.id)).toEqual([
      '201',
      '210',
      '202',
      '203',
    ]);
    expect(
      outer.elements.find((element) => element.id === '201'),
    ).toMatchObject({ left: 72, top: 54, width: 72, height: 54, rotate: 0 });
    expect(
      outer.elements.find((element) => element.id === '202'),
    ).toMatchObject({
      left: 108,
      top: 126,
      width: 216,
      height: 72,
      rotate: 90,
    });
    expect(
      outer.elements.find((element) => element.id === '203'),
    ).toMatchObject({
      left: -90,
      top: -18,
      width: 108,
      height: 144,
      rotate: 270,
    });

    const nestedResult = findGroup(outer.elements, '210');
    expect(nestedResult).toMatchObject({
      left: 72,
      top: 54,
      width: 144,
      height: 216,
    });
    expect(nestedResult.elements[0]).toMatchObject({
      id: '211',
      left: 0,
      top: 0,
      width: 72,
      height: 108,
    });
    const deeplyNestedResult = findGroup(nestedResult.elements, '212');
    expect(deeplyNestedResult).toMatchObject({
      left: 36,
      top: 54,
      width: 72,
      height: 108,
    });
    expect(deeplyNestedResult.elements).toEqual([
      expect.objectContaining({
        id: '213',
        left: 36,
        top: 27,
        width: 18,
        height: 54,
      }),
    ]);

    const uniformResult = findGroup(elements, '220');
    expect(uniformResult).toMatchObject({
      isFlipH: false,
      isFlipV: false,
      left: 0,
      top: 288,
      width: 144,
      height: 144,
    });
    expect(uniformResult.elements[0]).toMatchObject({
      id: '221',
      left: 0,
      top: 0,
      width: 72,
      height: 36,
      rotate: 90,
    });

    const zeroResult = findGroup(elements, '230');
    expect(zeroResult).toMatchObject({
      left: 216,
      top: 288,
      width: 72,
      height: 72,
    });
    expect(zeroResult.elements[0]).toMatchObject({
      id: '231',
      left: 0,
      top: 0,
      width: 0,
      height: 0,
    });
    expect(elements.some((element) => element.id === '240')).toBe(false);
  });
});
