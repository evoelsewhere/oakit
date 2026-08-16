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

const MARKUP_COMPATIBILITY_NS =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const OFFICE_MATH_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';

function connector(idAttribute: string, left: number): string {
  return `
    <p:cxnSp>
      <p:nvCxnSpPr><p:cNvPr ${idAttribute} name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
      <p:spPr>
        <a:xfrm><a:off x="${left}" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>
      </p:spPr>
    </p:cxnSp>`;
}

function mathAlternateContent(
  choiceId: string | undefined,
  fallbackId: string | undefined,
  relationshipId: string,
  expression: string,
  left: number,
  includeText = true,
): string {
  return `
    <mc:AlternateContent>
      <mc:Choice Requires="m">
        <p:sp>
          <p:nvSpPr><p:cNvPr ${choiceId ? `id="${choiceId}"` : ''} name="Equation"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="${left}" y="457200"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
          ${includeText ? `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Accessible ${expression}</a:t></a:r></a:p></p:txBody>` : ''}
        </p:sp>
        <m:oMath><m:r><m:t>${expression}</m:t></m:r></m:oMath>
      </mc:Choice>
      <mc:Fallback>
        <p:sp>
          <p:nvSpPr><p:cNvPr ${fallbackId ? `id="${fallbackId}"` : ''} name="Equation fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>
        </p:sp>
      </mc:Fallback>
    </mc:AlternateContent>`;
}

describe('PowerPoint element dispatch through the public API', () => {
  it('parses connectors, alternate groups, and math choice/fallback IDs', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:mc="${MARKUP_COMPATIBILITY_NS}" xmlns:m="${OFFICE_MATH_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${connector('id="110"', 0)}
            ${connector('', 914400)}
            <mc:AlternateContent>
              <mc:Fallback>
                <p:nvGrpSpPr><p:cNvPr id="120" name="Fallback group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="1828800" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="121" name="Fallback child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr>
                </p:sp>
              </mc:Fallback>
            </mc:AlternateContent>
            <mc:AlternateContent>
              <mc:Fallback>
                <p:nvGrpSpPr><p:cNvPr name="Anonymous fallback group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="2743200" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="123" name="Anonymous child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></p:spPr>
                </p:sp>
              </mc:Fallback>
            </mc:AlternateContent>
            ${mathAlternateContent('130', '131', 'rIdMathOne', 'x', 0)}
            ${mathAlternateContent(undefined, '140', 'rIdMathTwo', 'y', 914400)}
            ${mathAlternateContent(undefined, undefined, 'rIdMathThree', 'z', 1828800)}
            ${mathAlternateContent('150', '151', 'rIdMathFour', 'w', 2743200, false)}
            <mc:AlternateContent>
              <mc:Choice Requires="m">
                <p:sp><p:nvSpPr><p:cNvPr id="160" name="Not math"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
              </mc:Choice>
            </mc:AlternateContent>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdMathOne" Type="${OFFICE_REL_TYPE}image" Target="../media/math-one.png"/>
          <Relationship Id="rIdMathTwo" Type="${OFFICE_REL_TYPE}image" Target="../media/math-two.png"/>
          <Relationship Id="rIdMathThree" Type="${OFFICE_REL_TYPE}image" Target="../media/math-three.png"/>
          <Relationship Id="rIdMathFour" Type="${OFFICE_REL_TYPE}image" Target="../media/math-four.png"/>
        </Relationships>`,
    });

    const document = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const elements = document.slides[0]?.elements ?? [];

    expect(elements.find((element) => element.id === '110')).toMatchObject({
      type: 'shape',
      name: 'Connector',
      shapType: 'line',
    });
    expect(elements.find((element) => element.id === '')).toMatchObject({
      type: 'shape',
      name: 'Connector',
      shapType: 'line',
    });

    const fallbackGroup = elements.find((element) => element.id === '120');
    expect(fallbackGroup?.type).toBe('group');
    if (fallbackGroup?.type !== 'group') {
      throw new Error('Expected an alternate fallback group');
    }
    expect(fallbackGroup.elements).toContainEqual(
      expect.objectContaining({ id: '121', type: 'shape' }),
    );
    const anonymousGroup = elements.find(
      (element) => element.type === 'group' && element.id === '',
    );
    expect(anonymousGroup?.type).toBe('group');
    if (anonymousGroup?.type !== 'group') {
      throw new Error('Expected an anonymous alternate fallback group');
    }
    expect(anonymousGroup.elements).toContainEqual(
      expect.objectContaining({ id: '123', type: 'shape' }),
    );

    const firstMath = elements.find((element) => element.id === '130');
    const secondMath = elements.find((element) => element.id === '140');
    expect(firstMath?.type).toBe('math');
    expect(secondMath?.type).toBe('math');
    if (firstMath?.type !== 'math' || secondMath?.type !== 'math') {
      throw new Error('Expected both alternate math elements');
    }
    expect(firstMath).toMatchObject({
      height: 36,
      left: 0,
      type: 'math',
      latex: 'x',
      picRef: 'ppt/media/math-one.png',
      top: 36,
      width: 72,
    });
    expect(firstMath.text).toContain('Accessible&nbsp;x');
    expect(secondMath).toMatchObject({
      height: 36,
      left: 72,
      type: 'math',
      latex: 'y',
      picRef: 'ppt/media/math-two.png',
      top: 36,
      width: 72,
    });
    expect(secondMath.text).toContain('Accessible&nbsp;y');
    expect(
      elements.find(
        (element) => element.type === 'math' && element.latex === 'z',
      ),
    ).toMatchObject({
      id: '',
      left: 144,
      picRef: 'ppt/media/math-three.png',
      type: 'math',
    });
    expect(elements.find((element) => element.id === '150')).toMatchObject({
      id: '150',
      left: 216,
      latex: 'w',
      picRef: 'ppt/media/math-four.png',
      text: '',
      type: 'math',
    });
    expect(elements.some((element) => element.id === '160')).toBe(false);
  });
});
