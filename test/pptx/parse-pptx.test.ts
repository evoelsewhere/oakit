import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src/index';

async function createMinimalPptx(): Promise<Uint8Array> {
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Types>
        <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
      </Types>`,
  );
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation>
        <p:sldSz cx="9144000" cy="5143500"/>
      </p:presentation>`,
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      </Relationships>`,
  );
  zip.file(
    'ppt/theme/theme1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <a:theme>
        <a:themeElements>
          <a:clrScheme>
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
          </a:clrScheme>
          <a:fmtScheme/>
        </a:themeElements>
      </a:theme>`,
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="TextBox 1"/>
                <p:cNvSpPr/>
                <p:nvPr><p:ph type="body"/></p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="914400" y="914400"/>
                  <a:ext cx="1828800" cy="914400"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:noFill/>
                <a:ln><a:noFill/></a:ln>
              </p:spPr>
              <p:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:r>
                    <a:rPr lang="en-US" sz="1800"/>
                    <a:t>Hello AI</a:t>
                  </a:r>
                </a:p>
              </p:txBody>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>`,
  );
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      </Relationships>`,
  );
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:sldLayout><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
  );
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
      </Relationships>`,
  );
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
      <p:sldMaster>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMap/>
      </p:sldMaster>`,
  );
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
      </Relationships>`,
  );

  return zip.generateAsync({ type: 'uint8array' });
}

describe('parsePptx', () => {
  it('parses a minimal presentation package', async () => {
    const result = await parsePptx(await createMinimalPptx());

    expect(result.size).toEqual({ width: 720, height: 405 });
    expect(result.themeColors).toEqual([
      '#4472C4',
      '#ED7D31',
      '#A5A5A5',
      '#FFC000',
      '#5B9BD5',
      '#70AD47',
    ]);
    expect(result.usedFonts).toEqual([]);
    expect(result.slides).toHaveLength(1);
    expect(result.slides[0]).toMatchObject({
      fill: { type: 'color', value: '#fff' },
      layoutElements: [],
      note: '',
      transition: null,
    });
    expect(result.slides[0]?.elements).toHaveLength(1);
    expect(result.slides[0]?.elements[0]).toMatchObject({
      id: '2',
      type: 'text',
      left: 72,
      top: 72,
      width: 144,
      height: 72,
      name: 'TextBox 1',
    });
    expect(result.slides[0]?.elements[0]).toHaveProperty(
      'content',
      expect.stringContaining('Hello&nbsp;AI'),
    );
  });
});
