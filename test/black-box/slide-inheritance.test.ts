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
