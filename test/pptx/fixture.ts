import JSZip from 'jszip';

type PptxFixturePart = string | Uint8Array | null;

export type PptxFixtureOverrides = Record<string, PptxFixturePart>;

const MINIMAL_PPTX_PARTS: Readonly<Record<string, string>> = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
    <Types>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    </Types>`,
  'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation>
      <p:sldSz cx="9144000" cy="5143500"/>
    </p:presentation>`,
  'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    </Relationships>`,
  'ppt/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8"?>
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
  'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8"?>
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
  'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
  'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
  'ppt/slideMasters/slideMaster1.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldMaster>
      <p:cSld><p:spTree/></p:cSld>
      <p:clrMap/>
    </p:sldMaster>`,
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
    </Relationships>`,
};

export async function createMinimalPptx(
  overrides: PptxFixtureOverrides = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  const parts: Record<string, PptxFixturePart> = {
    ...MINIMAL_PPTX_PARTS,
    ...overrides,
  };

  for (const [filename, content] of Object.entries(parts)) {
    if (content !== null) zip.file(filename, content);
  }

  return zip.generateAsync({ type: 'uint8array' });
}
