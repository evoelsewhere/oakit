import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  independentTextSlide,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

const MIXED_ELEMENT_SLIDE = `<?xml version="1.0" encoding="UTF-8"?>
  <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:cSld>
      <p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Back"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Back</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="3" name="Middle"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
        </p:pic>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4" name="Front"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Front</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree>
    </p:cSld>
  </p:sld>`;

describe('PPTX public API adversarial cases', () => {
  it('preserves authored z-order across different element node types', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': MIXED_ELEMENT_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/image.png"/>
        </Relationships>`,
      'ppt/media/image.png': new Uint8Array([137, 80, 78, 71]),
    });

    const result = await parsePptx(input, { imageMode: 'none' });

    expect(result.slides[0]?.elements.map((element) => element.id)).toEqual([
      '2',
      '3',
      '4',
    ]);
  });

  it('accepts equivalent namespace prefixes in presentation XML', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `
        <q:presentation xmlns:q="${PRESENTATION_NS}" xmlns:rel="${OFFICE_REL_NS}">
          <q:sldIdLst><q:sldId id="256" rel:id="rIdSlide1"/></q:sldIdLst>
          <q:sldSz cx="9144000" cy="5143500"/>
        </q:presentation>`,
    });

    const result = await parsePptx(input);

    expect(result.slides).toHaveLength(1);
    expect(result.size).toEqual({ width: 720, height: 405 });
  });

  it('normalizes equivalent prefixes throughout DrawingML content', async () => {
    const slide = independentTextSlide('Aliased')
      .replace('xmlns:p=', 'xmlns:x=')
      .replace('xmlns:a=', 'xmlns:d=')
      .replaceAll('<p:', '<x:')
      .replaceAll('</p:', '</x:')
      .replaceAll('<a:', '<d:')
      .replaceAll('</a:', '</d:');
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': slide,
    });

    const result = await parsePptx(input);

    const element = result.slides[0]?.elements[0];
    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain('Aliased');
  });

  it('rejects mismatched XML closing names in strict mode', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `
        <p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}" xmlns:xp="urn:wrong">
          <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></xp:sldIdLst>
          <p:sldSz cx="9144000" cy="5143500"/>
        </p:presentation>`,
    });

    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'ppt/presentation.xml',
      },
    });
  });
});
