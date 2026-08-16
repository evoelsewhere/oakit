import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import { createMinimalPptx } from './fixture';

function textSlide(text: string): string {
  return `
    <p:sld>
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="${text}"/>
              <p:cNvSpPr txBox="1"/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="914400" cy="914400"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>${text}</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>`;
}

const SLIDE_RELATIONSHIPS = `
  <Relationships>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  </Relationships>`;

describe('PPTX slide order', () => {
  it('follows the presentation manifest and ignores orphan slide parts', async () => {
    const input = await createMinimalPptx({
      '[Content_Types].xml': `
        <Types>
          <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
          <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
          <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>`,
      'ppt/presentation.xml': `
        <p:presentation>
          <p:sldIdLst>
            <p:sldId id="256" r:id="rIdSlide2"/>
            <p:sldId id="257" r:id="rIdSlide1"/>
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="5143500"/>
        </p:presentation>`,
      'ppt/_rels/presentation.xml.rels': `
        <Relationships>
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
          <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
          <Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
        </Relationships>`,
      'ppt/slides/slide1.xml': textSlide('First'),
      'ppt/slides/slide2.xml': textSlide('Second'),
      'ppt/slides/slide3.xml': textSlide('Orphan'),
      'ppt/slides/_rels/slide1.xml.rels': SLIDE_RELATIONSHIPS,
      'ppt/slides/_rels/slide2.xml.rels': SLIDE_RELATIONSHIPS,
      'ppt/slides/_rels/slide3.xml.rels': SLIDE_RELATIONSHIPS,
    });

    const result = await parsePptx(input, { limits: { maxSlides: 2 } });

    expect(result.slides).toHaveLength(2);
    const second = result.slides[0]?.elements[0];
    const first = result.slides[1]?.elements[0];
    expect(second?.type).toBe('text');
    expect(first?.type).toBe('text');
    if (second?.type !== 'text' || first?.type !== 'text') {
      throw new Error('Expected ordered text elements');
    }
    expect(second.content).toContain('Second');
    expect(first.content).toContain('First');
  });

  it('enforces the configured slide count limit', async () => {
    const input = await createMinimalPptx({
      '[Content_Types].xml': `
        <Types>
          <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
          <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        </Types>`,
      'ppt/presentation.xml': `
        <p:presentation>
          <p:sldIdLst>
            <p:sldId id="256" r:id="rIdSlide1"/>
            <p:sldId id="257" r:id="rIdSlide2"/>
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="5143500"/>
        </p:presentation>`,
      'ppt/_rels/presentation.xml.rels': `
        <Relationships>
          <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
          <Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
        </Relationships>`,
      'ppt/slides/slide2.xml': textSlide('Second'),
      'ppt/slides/_rels/slide2.xml.rels': SLIDE_RELATIONSHIPS,
    });

    await expect(
      parsePptx(input, {
        limits: { maxSlides: 1 },
      }),
    ).rejects.toThrow('maxSlides');
  });

  it('does not select an external slide relationship in any TargetMode case', async () => {
    const input = await createMinimalPptx({
      'ppt/_rels/presentation.xml.rels': `
        <Relationships>
          <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="ppt/slides/slide1.xml" TargetMode="eXtErNaL"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides).toEqual([]);
  });
});
