import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  parsePptxRoundTripJson,
  parsePptxWithDiagnostics,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
  type PptxSceneDocument,
} from '../../src';
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

describe('PPTX public API in browsers', () => {
  it('renders self-contained SVG bytes that the browser can display', async () => {
    const bytes = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentTextSlide('Browser SVG preview'),
    });
    const inputBuffer = exactArrayBuffer(bytes);
    const result = await renderPptxToSvg(new Blob([inputBuffer]), {
      scale: 0.5,
      slideNumbers: [1],
    });
    const slide = result.slides[0];
    if (!slide) throw new Error('Expected one browser SVG slide');
    const source = new TextDecoder().decode(slide.data);
    const parsed = new DOMParser().parseFromString(source, 'image/svg+xml');

    expect(slide).toMatchObject({
      format: 'svg',
      height: 203,
      mimeType: 'image/svg+xml',
      slideNumber: 1,
      width: 360,
    });
    expect(parsed.querySelector('parsererror')).toBeNull();
    expect(parsed.querySelector('title')?.textContent).toBe(
      'PowerPoint slide 1',
    );
    expect(parsed.querySelector('tspan')?.textContent).toBe(
      'Browser\u00a0SVG\u00a0preview',
    );
    expect(source).not.toContain('<foreignObject');
    expect(source).not.toMatch(/(?:blob|file|https):/i);

    const url = URL.createObjectURL(
      new Blob([exactArrayBuffer(slide.data)], { type: slide.mimeType }),
    );
    try {
      const image = new Image();
      const loaded = new Promise<void>((resolve, reject) => {
        image.onload = () => resolve();
        image.onerror = () => reject(new Error('Browser could not load SVG'));
      });
      image.src = url;
      await loaded;
      expect(image.naturalWidth).toBe(360);
      expect(image.naturalHeight).toBe(203);
    } finally {
      URL.revokeObjectURL(url);
    }
  });

  it('preserves exact package bytes through a portable JSON agent hand-off', async () => {
    const bytes = await createIndependentPptx({
      'customXml/browser-agent.xml':
        '<?xml version="1.0"?><agent xmlns="urn:oakit:browser">preserve me</agent>',
      'ppt/slides/slide1.xml': independentTextSlide(
        'Portable browser hand-off',
      ),
    });
    const inputBuffer = exactArrayBuffer(bytes);
    const [byteSnapshot, blobSnapshot] = await Promise.all([
      readPptxRoundTrip(bytes),
      readPptxRoundTrip(
        new Blob([inputBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        }),
      ),
    ]);
    const [bytePortable, blobPortable] = await Promise.all([
      serializePptxRoundTripJson(byteSnapshot),
      serializePptxRoundTripJson(blobSnapshot),
    ]);
    const wireValue: unknown = JSON.parse(JSON.stringify(bytePortable));
    const restored = await parsePptxRoundTripJson(wireValue);
    const output = await writePptxRoundTrip(restored);
    const decoded = Uint8Array.from(
      atob(bytePortable.source.packageBase64),
      (character) => character.charCodeAt(0),
    );

    expect(blobSnapshot.source.data).toBeInstanceOf(Blob);
    expect(blobPortable).toEqual(bytePortable);
    expect(decoded).toEqual(bytes);
    expect(output.data).toEqual(bytes);
    expect(output.data).not.toBe(bytes);
    expect(output.report.level).toBe('R0');
  });

  it('creates deterministic packages and strictly reads them back', async () => {
    const scene: PptxSceneDocument = {
      layouts: [],
      masters: [],
      media: [],
      schemaVersion: 2,
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [
            {
              authored: {
                transform: { height: 60, width: 240, x: 20, y: 30 },
              },
              key: 'browser-text',
              resolved: { hidden: false },
              text: {
                body: {},
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'browser-run',
                        text: 'Created in browser',
                        type: 'run',
                      },
                    ],
                    key: 'browser-paragraph',
                  },
                ],
              },
              type: 'text',
            },
          ],
          key: 'browser-slide',
        },
      ],
      themes: [],
    };

    const [first, second] = await Promise.all([
      createPptx(scene),
      createPptx(scene),
    ]);
    const parsed = await parsePptx(first.data, {
      errorMode: 'strict',
      imageMode: 'none',
    });

    expect(second.data).toEqual(first.data);
    expect(first.report.level).toBe('C2');
    expect(first.report.addedPartCount).toBe(11);
    expect(parsed.size).toEqual({ height: 540, width: 960 });
    expect(parsed.slides).toHaveLength(1);
    expect(JSON.stringify(parsed)).toContain('Created&nbsp;in&nbsp;browser');
  });

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
