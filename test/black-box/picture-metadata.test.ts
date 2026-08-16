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

const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main';

function basicPicture(id: number, geometry = ''): string {
  return `<p:pic>
    <p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
    <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
      ${geometry}
    </p:spPr>
  </p:pic>`;
}

describe('PowerPoint picture metadata through the public API', () => {
  it('preserves transforms, inheritance, geometry, borders, filters, and safe links', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:a14="${A14_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="710" name="Rich picture"><a:hlinkClick r:id="rIdSafeLink"/></p:cNvPr>
                <p:cNvPicPr/><p:nvPr/>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rIdImage">
                  <a:extLst><a:ext uri="picture-effects">
                    <a14:imgProps><a14:imgLayer>
                      <a14:imgEffect>
                        <a14:saturation sat="125000"/>
                        <a14:brightnessContrast bright="-25000" contrast="50000"/>
                        <a14:sharpenSoften amount="40000"/>
                        <a14:colorTemperature colorTemp="6500"/>
                      </a14:imgEffect>
                      <a14:imgEffect><a14:sharpenSoften amount="-30000"/></a14:imgEffect>
                    </a14:imgLayer></a14:imgProps>
                  </a:ext></a:extLst>
                </a:blip>
                <a:stretch><a:fillRect/></a:stretch>
              </p:blipFill>
              <p:spPr>
                <a:xfrm flipV="1" flipH="1" rot="5400000">
                  <a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/>
                </a:xfrm>
                <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                <a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>
              </p:spPr>
            </p:pic>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="711" name="Inherited picture"/><p:cNvPicPr/>
                <p:nvPr><p:ph idx="42"/></p:nvPr>
              </p:nvPicPr>
              <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
              <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            </p:pic>
            <p:pic>
              <p:nvPicPr><p:cNvPr id="712" name="Missing image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
              <p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr>
            </p:pic>
            ${basicPicture(720)}
            ${basicPicture(721, '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>')}
            ${basicPicture(
              722,
              `<a:custGeom><a:avLst/><a:pathLst><a:path w="100" h="100">
                <a:moveTo><a:pt x="0" y="100"/></a:moveTo>
                <a:lnTo><a:pt x="50" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/>
              </a:path></a:pathLst></a:custGeom>`,
            )}
            ${basicPicture(723, '<a:custGeom><a:avLst/><a:pathLst/></a:custGeom>')}
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/picture.png"/>
          <Relationship Id="rIdSafeLink" Type="${OFFICE_REL_TYPE}hyperlink" Target=" https://example.com/picture " TargetMode="External"/>
        </Relationships>`,
      'ppt/slideLayouts/slideLayout1.xml': `
        <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="800" name="Picture position"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="42"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm flipH="1" rot="10800000"><a:off x="1828800" y="914400"/><a:ext cx="2743200" cy="1371600"/></a:xfrm></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const elements = result.slides[0]?.elements ?? [];
    const byId = Object.fromEntries(
      elements.map((element) => [element.id, element]),
    );

    expect(elements).toHaveLength(6);
    expect(byId['710']).toMatchObject({
      borderColor: '#123456',
      borderType: 'solid',
      borderWidth: 1,
      filters: {
        brightness: -0.25,
        colorTemperature: 6500,
        contrast: 0.5,
        saturation: 1.25,
        sharpen: 0.4,
        soften: 0.3,
      },
      geom: 'roundRect',
      height: 72,
      isFlipH: true,
      isFlipV: true,
      left: 72,
      link: 'https://example.com/picture',
      rotate: 90,
      top: 36,
      type: 'image',
      width: 144,
    });
    expect(byId['711']).toMatchObject({
      geom: 'rect',
      height: 108,
      isFlipH: true,
      isFlipV: false,
      left: 144,
      rotate: 180,
      top: 72,
      type: 'image',
      width: 216,
    });
    expect(byId).not.toHaveProperty('712');
    expect(byId['720']).toMatchObject({
      geom: 'rect',
      isFlipH: false,
      isFlipV: false,
      type: 'image',
    });
    expect(byId['720']).not.toHaveProperty('filters');
    expect(byId['720']).not.toHaveProperty('link');
    expect(byId['721']).toMatchObject({ geom: 'ellipse', type: 'image' });
    expect(byId['722']).toMatchObject({
      geom: 'custom:triangle',
      type: 'image',
    });
    expect(byId['723']).toMatchObject({ geom: 'custom', type: 'image' });
  });
});
