import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import { createMinimalPptx } from './fixture';

const IMAGE_SLIDE = `
  <p:sld>
    <p:cSld>
      <p:spTree>
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="2" name="Encoded path image"/>
            <p:cNvPicPr/>
            <p:nvPr/>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip r:embed="rIdImage"/>
            <a:stretch><a:fillRect/></a:stretch>
          </p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="914400" cy="914400"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
        </p:pic>
      </p:spTree>
    </p:cSld>
  </p:sld>`;

describe('PPTX media paths', () => {
  it('loads package media whose relationship target contains XML entities', async () => {
    const input = await createMinimalPptx({
      'ppt/slides/slide1.xml': IMAGE_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships>
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image&amp;one.png"/>
        </Relationships>`,
      'ppt/media/image&one.png': new Uint8Array([137, 80, 78, 71]),
    });

    const result = await parsePptx(input);
    const element = result.slides[0]?.elements[0];

    expect(element?.type).toBe('image');
    if (element?.type !== 'image') throw new Error('Expected an image element');
    expect(element.ref).toBe('ppt/media/image&one.png');
    expect(element.base64).toBe('data:image/png;base64,iVBORw==');
  });
});
