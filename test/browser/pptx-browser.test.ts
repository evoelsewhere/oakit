import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  independentTextSlide,
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

const MEDIA_SLIDE = `
  <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="Browser video"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rIdVideo"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:pic>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Browser audio"/><p:cNvPicPr/><p:nvPr><a:audioFile r:link="rIdAudio"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:pic>
    </p:spTree></p:cSld>
  </p:sld>`;

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}

async function createImagePackage(imageBytes: Uint8Array): Promise<Uint8Array> {
  return createIndependentPptx({
    'ppt/slides/slide1.xml': IMAGE_SLIDE,
    'ppt/slides/_rels/slide1.xml.rels': `
      <Relationships xmlns="${PACKAGE_REL_NS}">
        <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
        <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/image.png"/>
      </Relationships>`,
    'ppt/media/image.png': imageBytes,
  });
}

describe('PPTX public API in Chromium', () => {
  it('keeps input forms and concurrent parses deterministic', async () => {
    const [firstBytes, secondBytes] = await Promise.all([
      createIndependentPptx({
        'ppt/slides/slide1.xml': independentTextSlide('Browser first'),
      }),
      createIndependentPptx({
        'ppt/slides/slide1.xml': independentTextSlide('Browser second'),
      }),
    ]);
    const firstBuffer = exactArrayBuffer(firstBytes);

    const [fromBytes, fromBuffer, fromBlob, second, secondAgain] =
      await Promise.all([
        parsePptx(firstBytes),
        parsePptx(firstBuffer),
        parsePptx(new Blob([firstBuffer])),
        parsePptx(secondBytes),
        parsePptx(secondBytes),
      ]);

    expect(fromBuffer).toEqual(fromBytes);
    expect(fromBlob).toEqual(fromBytes);
    expect(secondAgain).toEqual(second);
    expect(JSON.stringify(fromBytes)).toContain('Browser&nbsp;first');
    expect(JSON.stringify(second)).toContain('Browser&nbsp;second');
    expect(second).not.toEqual(fromBytes);
  });

  it('reports malformed XML and enforces Blob input limits', async () => {
    const bytes = await createIndependentPptx({
      'ppt/theme/theme1.xml': '<a:theme><a:broken></a:theme>',
    });
    const inputBuffer = exactArrayBuffer(bytes);
    const tolerant = await parsePptxWithDiagnostics(new Blob([inputBuffer]));

    expect(tolerant.document.slides).toHaveLength(1);
    expect(tolerant.diagnostics).toHaveLength(1);
    expect(tolerant.diagnostics[0]).toMatchObject({
      code: 'xml-parse-failed',
      part: 'ppt/theme/theme1.xml',
      severity: 'warning',
    });
    await expect(
      parsePptx(bytes, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'ppt/theme/theme1.xml',
      },
    });
    await expect(
      parsePptx(new Blob([inputBuffer]), {
        limits: { maxInputBytes: bytes.byteLength - 1 },
      }),
    ).rejects.toMatchObject({
      diagnostic: { code: 'resource-limit-exceeded' },
    });
  });

  it('encodes embedded image bytes without Node buffer globals', async () => {
    const imageBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const document = await parsePptx(await createImagePackage(imageBytes), {
      errorMode: 'strict',
      imageMode: 'base64',
    });
    const image = document.slides[0]?.elements[0];
    if (image?.type !== 'image') throw new Error('Expected an image element');

    expect(image.base64).toBe('data:image/png;base64,iVBORw0KGgo=');
    expect(
      Uint8Array.from(atob(image.base64.split(',')[1] ?? ''), (character) =>
        character.charCodeAt(0),
      ),
    ).toEqual(imageBytes);
  });

  it('accepts a browser Blob and exposes embedded image bytes as a blob URL', async () => {
    const imageBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const bytes = await createImagePackage(imageBytes);
    const inputBuffer = exactArrayBuffer(bytes);

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

  it('exposes exact embedded audio and video bytes as browser blob URLs', async () => {
    const posterBytes = new Uint8Array([137, 80, 78, 71]);
    const videoBytes = new Uint8Array([0, 0, 0, 24, 102, 116, 121, 112]);
    const audioBytes = new Uint8Array([73, 68, 51, 4, 0, 0, 0, 0]);
    const bytes = await createIndependentPptx({
      'ppt/slides/slide1.xml': MEDIA_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdPoster" Type="${OFFICE_REL_TYPE}image" Target="../media/poster.png"/>
          <Relationship Id="rIdVideo" Type="${OFFICE_REL_TYPE}video" Target="../media/clip.mp4"/>
          <Relationship Id="rIdAudio" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.mp3"/>
        </Relationships>`,
      'ppt/media/poster.png': posterBytes,
      'ppt/media/clip.mp4': videoBytes,
      'ppt/media/sound.mp3': audioBytes,
    });
    const document = await parsePptx(bytes, {
      audioMode: 'blob',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'blob',
    });
    const video = document.slides[0]?.elements.find(
      (element) => element.type === 'video',
    );
    const audio = document.slides[0]?.elements.find(
      (element) => element.type === 'audio',
    );
    if (video?.type !== 'video' || audio?.type !== 'audio') {
      throw new Error('Expected browser media elements');
    }

    try {
      const [loadedVideo, loadedAudio] = await Promise.all([
        fetch(video.blob).then(async (response) =>
          response.arrayBuffer().then((buffer) => new Uint8Array(buffer)),
        ),
        fetch(audio.blob).then(async (response) =>
          response.arrayBuffer().then((buffer) => new Uint8Array(buffer)),
        ),
      ]);
      expect(loadedVideo).toEqual(videoBytes);
      expect(loadedAudio).toEqual(audioBytes);
    } finally {
      URL.revokeObjectURL(video.blob);
      URL.revokeObjectURL(audio.blob);
    }
  });
});
