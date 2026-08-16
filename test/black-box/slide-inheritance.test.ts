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

function themeWithMinorFont(name: string, typeface: string): string {
  return `
    <a:theme xmlns:a="${DRAWING_NS}" name="${name}">
      <a:themeElements>
        <a:clrScheme name="${name}">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
        <a:fontScheme name="${name}">
          <a:majorFont><a:latin typeface="${typeface} Major"/></a:majorFont>
          <a:minorFont><a:latin typeface="${typeface}"/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="${name}"/>
      </a:themeElements>
    </a:theme>`;
}

function layoutWithTransition(transition: string): string {
  return `
    <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
      ${transition}
    </p:sldLayout>`;
}

function masterWithTransition(transition: string): string {
  return `
    <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
      <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
      ${transition}
    </p:sldMaster>`;
}

describe('PowerPoint slide inheritance through the public API', () => {
  it('parses slide content without an optional layout relationship', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}"/>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]).toMatchObject({
      fill: { type: 'color', value: '#fff' },
      layoutElements: [],
      note: '',
      transition: null,
    });
    expect(result.slides[0]?.elements).toHaveLength(1);
    expect(result.slides[0]?.elements[0]?.type).toBe('text');
  });

  it('resolves picture resources owned by the slide layout', async () => {
    const input = await createIndependentPptx({
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr/><p:grpSpPr/>
            <p:pic>
              <p:nvPicPr><p:cNvPr id="30" name="Layout badge"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
              <p:blipFill><a:blip r:embed="rIdLayoutImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm></p:spPr>
            </p:pic>
          </p:spTree></p:cSld>
        </p:sldLayout>`,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdMaster" Type="${OFFICE_REL_TYPE}slideMaster" Target="../slideMasters/slideMaster1.xml"/>
          <Relationship Id="rIdLayoutImage" Type="${OFFICE_REL_TYPE}image" Target="../media/layout-badge.png"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });

    expect(result.slides[0]?.layoutElements).toContainEqual(
      expect.objectContaining({
        type: 'image',
        id: '30',
        ref: 'ppt/media/layout-badge.png',
      }),
    );
  });

  it('inherits freeform text styling from the slide master', async () => {
    const input = await createIndependentPptx({
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
          <p:txStyles>
            <p:otherStyle>
              <a:lvl1pPr><a:defRPr sz="3100"><a:solidFill><a:srgbClr val="A12345"/></a:solidFill></a:defRPr></a:lvl1pPr>
            </p:otherStyle>
          </p:txStyles>
        </p:sldMaster>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const element = result.slides[0]?.elements[0];

    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain('font-size: 31pt;');
    expect(element.content).toContain('color: #A12345;');
  });

  it('applies the referenced package-level table style', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="40" name="Styled table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                  <a:tblPr><a:tableStyleId>{OAKIT-STYLE}</a:tableStyleId></a:tblPr>
                  <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                  <a:tr h="457200"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Styled</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                </a:tbl>
              </a:graphicData></a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/tableStyles.xml': `
        <a:tblStyleLst xmlns:a="${DRAWING_NS}">
          <a:tblStyle styleId="{OAKIT-STYLE}" styleName="OAKit">
            <a:wholeTbl>
              <a:tcTxStyle b="1"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr><a:top><a:ln w="25400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></a:top></a:tcBdr>
                <a:fill><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></a:fill>
              </a:tcStyle>
            </a:wholeTbl>
          </a:tblStyle>
        </a:tblStyleLst>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const table = result.slides[0]?.elements[0];

    expect(table?.type).toBe('table');
    if (table?.type !== 'table') throw new Error('Expected a table element');
    expect(table.borders.top).toMatchObject({
      borderColor: '#112233',
      borderWidth: 2,
    });
    expect(table.data[0]?.[0]).toMatchObject({
      fillColor: '#DDEEFF',
      fontBold: true,
      fontColor: '#445566',
    });
  });

  it('resolves picture fills owned directly by the slide', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="50" name="Photo fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:blipFill><a:blip r:embed="rIdSlideFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
              </p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdSlideFill" Type="${OFFICE_REL_TYPE}image" Target="../media/slide-fill.png"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });

    expect(result.slides[0]?.elements[0]).toMatchObject({
      type: 'shape',
      id: '50',
      fill: {
        type: 'image',
        value: {
          ref: 'ppt/media/slide-fill.png',
          base64: '',
          blob: '',
          opacity: 1,
        },
      },
    });
  });

  it('accepts only safe shape-level hyperlink relationships', async () => {
    const linkedShape = (id: string, relationshipId: string, left: number) => `
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="${id}" name="Linked ${id}"><a:hlinkClick r:id="${relationshipId}"/></p:cNvPr>
          <p:cNvSpPr/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${left}" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
        </p:spPr>
      </p:sp>`;
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${linkedShape('60', 'rIdSafeShape', 0)}
            ${linkedShape('61', 'rIdWrongType', 457200)}
            ${linkedShape('62', 'rIdUnsafeShape', 914400)}
            ${linkedShape('63', 'rIdMissingShape', 1371600)}
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdSafeShape" Type="${OFFICE_REL_TYPE}hyperlink" Target="https://example.com/shape?a=1&amp;b=2" TargetMode="External"/>
          <Relationship Id="rIdWrongType" Type="${OFFICE_REL_TYPE}image" Target="https://example.com/not-a-hyperlink" TargetMode="External"/>
          <Relationship Id="rIdUnsafeShape" Type="${OFFICE_REL_TYPE}hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(byId['60']).toMatchObject({
      id: '60',
      link: 'https://example.com/shape?a=1&b=2',
    });
    expect(byId['61']).not.toHaveProperty('link');
    expect(byId['62']).not.toHaveProperty('link');
    expect(byId['63']).not.toHaveProperty('link');
  });

  it('matches inherited placeholders by index before their shared type', async () => {
    const placeholder = (
      id: string,
      index: string,
      left: number,
      width: number,
      fill = '',
    ) => `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${index}"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="${index}"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${left}" y="0"/><a:ext cx="${width}" cy="914400"/></a:xfrm>
          ${fill}
        </p:spPr>
      </p:sp>`;
    const slidePlaceholder = (id: string, index: string | undefined) => `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="${id}" name="Slide ${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="body"${index ? ` idx="${index}"` : ''}/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide ${id}</a:t></a:r></a:p></p:txBody>
      </p:sp>`;
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr/><p:grpSpPr/>
            ${slidePlaceholder('80', '7')}
            ${slidePlaceholder('81', undefined)}
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            ${placeholder('70', '8', 0, 457200)}
            ${placeholder('71', '7', 914400, 1828800)}
          </p:spTree></p:cSld>
        </p:sldLayout>`,
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            ${placeholder('72', '8', 2743200, 2743200, '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>')}
            ${placeholder('73', '7', 3657600, 3657600, '<a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill>')}
          </p:spTree></p:cSld>
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const indexed = result.slides[0]?.elements.find(
      (element) => element.id === '80',
    );
    const typed = result.slides[0]?.elements.find(
      (element) => element.id === '81',
    );

    expect(indexed).toMatchObject({
      left: 72,
      width: 144,
      fill: { type: 'color', value: '#AABBCC' },
    });
    expect(typed).toMatchObject({
      left: 0,
      width: 36,
      fill: { type: 'color', value: '#112233' },
    });
    expect(result.slides[0]?.layoutElements).toEqual([]);
  });

  it('hides master shapes when the layout disables them', async () => {
    const input = await createIndependentPptx({
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" showMasterSp="0">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="90" name="Layout shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>`,
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="91" name="Hidden master shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="457200" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="654321"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(
      result.slides[0]?.layoutElements.map((element) => element.id),
    ).toEqual(['90']);
  });

  it('uses the slide master theme before the presentation theme', async () => {
    const input = await createIndependentPptx({
      'ppt/theme/theme1.xml': themeWithMinorFont(
        'Presentation Theme',
        'Presentation Font',
      ),
      'ppt/theme/theme2.xml': themeWithMinorFont('Master Theme', 'Master Font'),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="../theme/theme2.xml"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const element = result.slides[0]?.elements[0];

    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain('font-family: &quot;Master Font&quot;;');
    expect(element.content).not.toContain('Presentation Font');
  });

  it('falls back to the presentation theme when the master has none', async () => {
    const input = await createIndependentPptx({
      'ppt/theme/theme1.xml': themeWithMinorFont(
        'Presentation Theme',
        'Presentation Font',
      ),
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}"/>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const element = result.slides[0]?.elements[0];

    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain(
      'font-family: &quot;Presentation Font&quot;;',
    );
  });

  it('resolves picture resources owned by the master and its theme', async () => {
    const input = await createIndependentPptx({
      'ppt/slideMasters/slideMaster1.xml': `
        <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld>
            <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
            <p:spTree>
              <p:nvGrpSpPr/><p:grpSpPr/>
              <p:pic>
                <p:nvPicPr><p:cNvPr id="20" name="Master logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                <p:blipFill><a:blip r:embed="rIdMasterImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
              </p:pic>
            </p:spTree>
          </p:cSld>
          <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
        </p:sldMaster>`,
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="../theme/theme2.xml"/>
          <Relationship Id="rIdMasterImage" Type="${OFFICE_REL_TYPE}image" Target="../media/master-logo.png"/>
        </Relationships>`,
      'ppt/theme/theme2.xml': `
        <a:theme xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}" name="Picture Theme">
          <a:themeElements>
            <a:clrScheme name="Picture"><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
            <a:fontScheme name="Picture"><a:majorFont/><a:minorFont/></a:fontScheme>
            <a:fmtScheme name="Picture">
              <a:bgFillStyleLst>
                <a:blipFill><a:blip r:embed="rIdThemeImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
              </a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>`,
      'ppt/theme/_rels/theme2.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdThemeImage" Type="${OFFICE_REL_TYPE}image" Target="../media/theme-background.png"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const masterImage = result.slides[0]?.layoutElements.find(
      (element) => element.type === 'image',
    );

    expect(masterImage).toMatchObject({
      type: 'image',
      id: '20',
      ref: 'ppt/media/master-logo.png',
    });
    expect(result.slides[0]?.fill).toEqual({
      type: 'image',
      value: {
        ref: 'ppt/media/theme-background.png',
        base64: '',
        blob: '',
        opacity: 1,
      },
    });
  });

  it('uses the layout transition before the master transition', async () => {
    const input = await createIndependentPptx({
      'ppt/slideLayouts/slideLayout1.xml': layoutWithTransition(
        '<p:transition spd="fast"><p:push dir="l"/></p:transition>',
      ),
      'ppt/slideMasters/slideMaster1.xml': masterWithTransition(
        '<p:transition spd="med"><p:wipe dir="r"/></p:transition>',
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]?.transition).toEqual({
      type: 'push',
      duration: 500,
      direction: 'l',
    });
  });

  it('uses the master transition when slide and layout omit one', async () => {
    const input = await createIndependentPptx({
      'ppt/slideMasters/slideMaster1.xml': masterWithTransition(
        '<p:transition advClick="0" advTm="321"><p:cover dir="d"/></p:transition>',
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]?.transition).toEqual({
      type: 'cover',
      duration: 1000,
      direction: 'd',
      autoNextAfter: 321,
    });
  });
});
