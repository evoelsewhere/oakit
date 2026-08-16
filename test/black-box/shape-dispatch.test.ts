import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

describe('PowerPoint shape dispatch through the public API', () => {
  it('resolves authored, indexed, inherited, and connector shape semantics exactly', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="300"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm flipH="1" flipV="1" rot="5400000"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst><a:gd name="valid" fmla="val 25000"/><a:gd name="trimmed" fmla=" val 10000 "/><a:gd name="multipleSpaces" fmla="val   15000"/><a:gd name="positive" fmla="val +5000"/><a:gd name="negative" fmla="val -5000"/><a:gd name="malformed" fmla="val nope"/><a:gd name="prefixed" fmla="junk val 10000"/><a:gd name="suffixed" fmla="val 10000 junk"/><a:gd name="characterPrefix" fmla="val x12000"/><a:gd name="calculated" fmla="*/ w 1 2"/><a:gd fmla="val 50000"/></a:avLst></a:prstGeom></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="301" name="Authored text box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="457200"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
              <p:txXfrm rot="5400000"/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Authored text</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="302" name="Layout typed"/><p:cNvSpPr/><p:nvPr><p:ph idx="7"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="303" name="Master typed"/><p:cNvSpPr/><p:nvPr><p:ph idx="8"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="304" name="Centered title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Centered title</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:cxnSp>
              <p:nvCxnSpPr><p:cNvPr id="305"/><p:cNvCxnSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvCxnSpPr>
              <p:spPr><a:xfrm><a:off x="4572000" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
            </p:cxnSp>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="400" name="Layout body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="401" name="Untyped layout"/><p:cNvSpPr/><p:nvPr><p:ph idx="8"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>`,
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="500" name="Master body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="8"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="2743200" y="914400"/><a:ext cx="1828800" cy="457200"/></a:xfrm></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="501" name="Master title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="3657600" y="1828800"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    const adjustedShape = byId['300'];
    expect(adjustedShape).toMatchObject({
      height: 36,
      content: '',
      isFlipH: true,
      isFlipV: true,
      left: 0,
      name: '',
      rotate: 90,
      top: 0,
      type: 'shape',
      width: 72,
    });
    expect(adjustedShape?.type).toBe('shape');
    if (adjustedShape?.type !== 'shape') {
      throw new Error('Expected the adjusted rectangle to remain a shape');
    }
    expect(adjustedShape.keypoints).toEqual({
      multipleSpaces: 0.3,
      negative: -0.1,
      positive: 0.1,
      trimmed: 0.2,
      valid: 0.5,
    });
    const authoredTextBox = byId['301'];
    expect(authoredTextBox).toMatchObject({
      height: 36,
      left: 0,
      name: 'Authored text box',
      isFlipH: false,
      isFlipV: false,
      rotate: 180,
      top: 36,
      type: 'text',
      width: 72,
    });
    expect(authoredTextBox?.type).toBe('text');
    if (authoredTextBox?.type !== 'text') {
      throw new Error('Expected the authored text box to remain text');
    }
    expect(authoredTextBox.content).toContain('Authored&nbsp;text');
    expect(byId['302']).toMatchObject({
      height: 36,
      left: 72,
      top: 72,
      type: 'text',
      width: 72,
    });
    expect(byId['303']).toMatchObject({
      height: 36,
      left: 144,
      top: 72,
      type: 'text',
      width: 72,
    });
    expect(byId['304']).toMatchObject({
      fill: { type: 'color', value: '#AABBCC' },
      height: 72,
      left: 288,
      top: 144,
      type: 'text',
      width: 144,
    });
    expect(byId['305']).toMatchObject({
      height: 36,
      left: 360,
      name: '',
      shapType: 'line',
      top: 0,
      type: 'shape',
      width: 72,
    });
  });
});
