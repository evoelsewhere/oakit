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

function textFlowSlide(): string {
  return `
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Text flow"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
              <a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:rPr/><a:t>One</a:t></a:fld>
              <a:fld id="{00000000-0000-0000-0000-000000000002}" type="slidenum"><a:rPr/><a:t>Two</a:t></a:fld>
            </a:p>
            <a:p>
              <a:r><a:rPr/><a:t>Before &amp; after</a:t></a:r>
              <a:br><a:rPr/></a:br>
              <a:r><a:rPr><a:hlinkClick r:id="rIdSafe"/></a:rPr><a:t>Safe</a:t></a:r>
              <a:br><a:rPr/></a:br>
              <a:r><a:rPr><a:hlinkClick r:id="rIdUnsafe"/></a:rPr><a:t>Unsafe</a:t></a:r>
            </a:p>
            <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:rPr/><a:t>Parent</a:t></a:r></a:p>
            <a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr><a:r><a:rPr/><a:t>Child</a:t></a:r></a:p>
            <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:rPr/><a:t>Sibling</a:t></a:r></a:p>
            <a:p><a:pPr lvl="1junk"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr/><a:t>Normalized</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`;
}

function slideRelationships(): string {
  return `
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rIdSafe" Type="${OFFICE_REL_TYPE}hyperlink" Target="https://example.com/path?a=1&amp;b=2" TargetMode="External"/>
      <Relationship Id="rIdUnsafe" Type="${OFFICE_REL_TYPE}hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
    </Relationships>`;
}

describe('PowerPoint text flow through the public API', () => {
  it('preserves mixed text semantics and emits valid nested list structure', async () => {
    const document = await parsePptx(
      await createIndependentPptx({
        'ppt/slides/slide1.xml': textFlowSlide(),
        'ppt/slides/_rels/slide1.xml.rels': slideRelationships(),
      }),
      { errorMode: 'strict' },
    );

    const element = document.slides[0]?.elements[0];
    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');

    expect(element.content).toContain('OneTwo');
    expect(element.content).toContain('Before&nbsp;&amp;&nbsp;after');
    expect(element.content.match(/<br>/g)).toHaveLength(2);
    expect(element.content).toContain(
      '<a href="https://example.com/path?a=1&amp;b=2" target="_blank" rel="noopener noreferrer">Safe</a>',
    );
    expect(element.content).not.toContain('javascript:');
    expect(element.content).toContain('>Unsafe</span>');
    expect(element.content).toMatch(
      /<ul><li><p[^>]*>.*Parent.*<\/p><ul><li><p[^>]*>.*Child.*<\/p><\/li><\/ul><\/li><li><p[^>]*>.*Sibling.*<\/p><\/li><\/ul><ol><li>/,
    );
    expect(element.content).not.toMatch(/NaN|Infinity/);
  });
});
