import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import { createMinimalPptx } from './fixture';

const MALICIOUS_TEXT_SLIDE = `
  <p:sld>
    <p:cSld>
      <p:spTree>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Unsafe text"/>
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
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr>
                <a:t><![CDATA[<img src=x onerror=alert(1)> & unsafe]]></a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:spTree>
    </p:cSld>
  </p:sld>`;

describe('PPTX HTML security', () => {
  it('escapes rich text and drops unsafe run hyperlinks', async () => {
    const input = await createMinimalPptx({
      'ppt/slides/slide1.xml': MALICIOUS_TEXT_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships>
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
        </Relationships>`,
    });

    const result = await parsePptx(input);
    const element = result.slides[0]?.elements[0];

    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain(
      '&lt;img&nbsp;src=x&nbsp;onerror=alert(1)&gt;&nbsp;&amp;&nbsp;unsafe',
    );
    expect(element.content).not.toContain('<img');
    expect(element.content).not.toContain('href=');
  });
});
