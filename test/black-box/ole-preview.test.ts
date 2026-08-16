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

const MARKUP_COMPATIBILITY_NS =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const POWERPOINT_2010_NS =
  'http://schemas.microsoft.com/office/powerpoint/2010/main';
const OLE_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole';

function previewPicture(id: number, relationshipId: string): string {
  return `<p:pic>
    <p:nvPicPr><p:cNvPr id="${id}" name="OLE preview ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr>
      <a:xfrm><a:off x="${id * 1000}" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    </p:spPr>
  </p:pic>`;
}

function graphicFrame(id: number, graphicData: string): string {
  return `<p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="${id}" name="OLE frame ${id}"/><p:cNvGraphicFramePr/><p:nvPr/>
    </p:nvGraphicFramePr>
    <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></p:xfrm>
    <a:graphic><a:graphicData uri="${OLE_URI}">${graphicData}</a:graphicData></a:graphic>
  </p:graphicFrame>`;
}

describe('PowerPoint OLE previews through the public API', () => {
  it('uses Choice before Fallback and also supports fallback-only and direct previews', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:mc="${MARKUP_COMPATIBILITY_NS}" xmlns:p14="${POWERPOINT_2010_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${graphicFrame(
              800,
              `<mc:AlternateContent>
                <mc:Choice Requires="p14"><p:oleObj>${previewPicture(900, 'rIdChoice')}</p:oleObj></mc:Choice>
                <mc:Fallback><p:oleObj>${previewPicture(901, 'rIdFallback')}</p:oleObj></mc:Fallback>
              </mc:AlternateContent>`,
            )}
            ${graphicFrame(
              801,
              `<mc:AlternateContent>
                <mc:Fallback><p:oleObj>${previewPicture(902, 'rIdFallbackOnly')}</p:oleObj></mc:Fallback>
              </mc:AlternateContent>`,
            )}
            ${graphicFrame(
              802,
              `<p:oleObj>${previewPicture(903, 'rIdDirect')}</p:oleObj>`,
            )}
            ${graphicFrame(803, '<mc:AlternateContent/>')}
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="804" name="Unknown"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <a:graphic><a:graphicData uri="urn:unsupported"/></a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdChoice" Type="${OFFICE_REL_TYPE}image" Target="../media/choice.png"/>
          <Relationship Id="rIdFallback" Type="${OFFICE_REL_TYPE}image" Target="../media/fallback.png"/>
          <Relationship Id="rIdFallbackOnly" Type="${OFFICE_REL_TYPE}image" Target="../media/fallback-only.png"/>
          <Relationship Id="rIdDirect" Type="${OFFICE_REL_TYPE}image" Target="../media/direct.png"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(Object.keys(byId)).toEqual(['800', '801', '802']);
    expect(byId['800']).toMatchObject({
      ref: 'ppt/media/choice.png',
      type: 'image',
    });
    expect(byId['801']).toMatchObject({
      ref: 'ppt/media/fallback-only.png',
      type: 'image',
    });
    expect(byId['802']).toMatchObject({
      ref: 'ppt/media/direct.png',
      type: 'image',
    });
  });
});
