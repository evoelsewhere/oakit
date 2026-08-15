import JSZip from 'jszip';

export type BlackBoxPart = string | Uint8Array | null;
export type BlackBoxOverrides = Record<string, BlackBoxPart>;

export const PRESENTATION_NS =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
export const DRAWING_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
export const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PACKAGE_REL_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
export const OFFICE_REL_TYPE = `${OFFICE_REL_NS}/`;

export function independentTextSlide(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Text"/>
              <p:cNvSpPr txBox="1"/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>`;
}

const BASE_PARTS: Readonly<Record<string, string>> = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
      <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
      <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    </Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rId1" Type="${OFFICE_REL_TYPE}officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`,
  'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="5143500"/>
    </p:presentation>`,
  'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="theme/theme1.xml"/>
      <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="slides/slide1.xml"/>
    </Relationships>`,
  'ppt/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <a:theme xmlns:a="${DRAWING_NS}" name="Independent Theme">
      <a:themeElements>
        <a:clrScheme name="Independent">
          <a:dk1><a:srgbClr val="000000"/></a:dk1>
          <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
        <a:fontScheme name="Independent">
          <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
          <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Independent"/>
      </a:themeElements>
    </a:theme>`,
  'ppt/slides/slide1.xml': independentTextSlide('Black box'),
  'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`,
  'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
    </p:sldLayout>`,
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdMaster" Type="${OFFICE_REL_TYPE}slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
  'ppt/slideMasters/slideMaster1.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
      <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
    </p:sldMaster>`,
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${PACKAGE_REL_NS}">
      <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="../theme/theme1.xml"/>
    </Relationships>`,
};

export async function createIndependentPptx(
  overrides: BlackBoxOverrides = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  const parts: Record<string, BlackBoxPart> = {
    ...BASE_PARTS,
    ...overrides,
  };

  for (const [name, content] of Object.entries(parts)) {
    if (content !== null) zip.file(name, content);
  }
  return zip.generateAsync({ compression: 'DEFLATE', type: 'uint8array' });
}
