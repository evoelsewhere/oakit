import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

describe('PowerPoint speaker notes through the public API', () => {
  it('renders exact safe text, alignment, fields, breaks, and nested lists', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdNotes" Type="${OFFICE_REL_TYPE}notesSlide" Target="../notesSlides/notesSlide1.xml"/>
        </Relationships>`,
      'ppt/notesSlides/notesSlide1.xml': `
        <p:notes xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="1" name="Ignored"/><p:nvPr><p:ph type="hdr"/></p:nvPr></p:nvSpPr>
              <p:txBody><a:p><a:r><a:t>Ignored header</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
              <p:txBody>
                <a:p><a:pPr algn="r"/><a:r><a:t><![CDATA[Right <&]]></a:t></a:r></a:p>
                <a:p><a:pPr algn="ctr"/><a:r><a:t>Center</a:t></a:r></a:p>
                <a:p><a:pPr algn="just"/><a:r><a:t>Justified</a:t></a:r></a:p>
                <a:p><a:pPr algn="dist"/><a:r><a:t>Distributed</a:t></a:r></a:p>
                <a:p><a:pPr algn="unknown"/><a:r><a:t>Default</a:t></a:r></a:p>
                <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:t>Parent</a:t></a:r></a:p>
                <a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr><a:r><a:t>Child</a:t></a:r></a:p>
                <a:p><a:pPr lvl="2"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Grandchild</a:t></a:r></a:p>
                <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:t>Sibling</a:t></a:r></a:p>
                <a:p><a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Numbered</a:t></a:r></a:p>
                <a:p><a:r><a:t>A</a:t></a:r><a:br/><a:fld id="field"><a:t>B&#x9;C</a:t></a:fld><a:r/></a:p>
                <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:t>End parent</a:t></a:r></a:p>
                <a:p><a:pPr lvl="1"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>End child</a:t></a:r></a:p>
              </p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:notes>`,
    });

    const document = await parsePptx(input, { errorMode: 'strict' });

    expect(document.slides[0]?.note).toBe(
      '<p style="text-align:right;">Right&nbsp;&lt;&amp;</p>' +
        '<p style="text-align:center;">Center</p>' +
        '<p style="text-align:justify;">Justified</p>' +
        '<p style="text-align:justify;">Distributed</p>' +
        '<p style="text-align:left;">Default</p>' +
        '<ul><li><p style="text-align:left;">Parent</p>' +
        '<ul><li><p style="text-align:left;">Child</p>' +
        '<ol><li><p style="text-align:left;">Grandchild</p></li></ol></li></ul>' +
        '</li><li><p style="text-align:left;">Sibling</p></li></ul>' +
        '<ol><li><p style="text-align:left;">Numbered</p></li></ol>' +
        '<p style="text-align:left;">A<br>B&nbsp;&nbsp;&nbsp;&nbsp;C</p>' +
        '<ul><li><p style="text-align:left;">End&nbsp;parent</p>' +
        '<ol><li><p style="text-align:left;">End&nbsp;child</p></li></ol></li></ul>',
    );
    expect(document.slides[0]?.note).not.toContain('Ignored header');
  });
});
