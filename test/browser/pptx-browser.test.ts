import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from '../black-box/pptx-package';

const IMAGE_SLIDE = `
  <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="Browser image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:pic>
    </p:spTree></p:cSld>
  </p:sld>`;

describe('PPTX public API in Chromium', () => {
  it('accepts a browser Blob and exposes embedded image bytes as a blob URL', async () => {
    const imageBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const bytes = await createIndependentPptx({
      'ppt/slides/slide1.xml': IMAGE_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/image.png"/>
        </Relationships>`,
      'ppt/media/image.png': imageBytes,
    });
    const inputBuffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;

    const document = await parsePptx(
      new Blob([inputBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }),
      { errorMode: 'strict', imageMode: 'blob' },
    );
    const image = document.slides[0]?.elements.find(
      (element) => element.type === 'image',
    );
    if (image?.type !== 'image') throw new Error('Expected an image element');

    try {
      const loadedBuffer: ArrayBuffer = await fetch(image.blob).then(
        (response) => response.arrayBuffer(),
      );
      const loadedBytes = new Uint8Array(loadedBuffer);
      expect(loadedBytes).toEqual(imageBytes);
      expect(navigator.userAgent).toContain('Chrome');
    } finally {
      URL.revokeObjectURL(image.blob);
    }
  });
});
