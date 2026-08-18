import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { decodeBase64 } from '../../src/common/binary/base64';
import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  setPptxRoundTripImageTransform,
  setPptxRoundTripShapeTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
  type PptxSceneTransform,
} from '../../src';

const PNG_BYTES = decodeBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

function scene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [{ data: PNG_BYTES, key: 'source-media', mimeType: 'image/png' }],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 90, width: 120, x: 500, y: 300 },
            },
            key: 'source-picture',
            mediaKey: 'source-media',
            resolved: { hidden: false },
            type: 'image',
          },
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

const CHANGED_TRANSFORM: PptxSceneTransform = {
  flipHorizontal: true,
  flipVertical: true,
  height: 130,
  rotation: 15,
  width: 170,
  x: 450,
  y: 260,
};

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint image transform editing', () => {
  it('patches only picture geometry and preserves exact media bytes', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    expect(snapshot.document.slides[0]?.elements[0]).toMatchObject({
      key: 'slide-1-element-1',
      type: 'image',
    });

    const edited = await setPptxRoundTripImageTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts, verified, rendered] = await Promise.all([
      payloads(created.data),
      payloads(output.data),
      readPptxRoundTrip(output.data),
      renderPptxToSvg(output.data, { slideNumbers: [1] }),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      resolved: { transform: CHANGED_TRANSFORM },
      type: 'image',
    });
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') {
        expect(result).not.toEqual(source);
      } else {
        expect(result, name).toEqual(source);
      }
    }
    expect(outputParts.get('ppt/media/image1.png')).toEqual(PNG_BYTES);
    expect(new TextDecoder().decode(rendered.slides[0]?.data)).toContain(
      'data:image/png;base64,',
    );
  });

  it('survives portable JSON with native capability binding', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const edited = await setPptxRoundTripImageTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });
    const portable = await serializePptxRoundTripJson(edited);
    const restored = await parsePptxRoundTripJson(
      JSON.parse(JSON.stringify(portable)),
    );
    const output = await writePptxRoundTrip(restored);

    expect(restored.consistency.capabilityProfileVersion).toBe(
      'pptx-roundtrip-native-v1',
    );
    expect(output.report.operations).toMatchObject([
      { kind: 'set-transform', status: 'verified' },
    ]);
  });

  it('keeps image and shape transform APIs type-specific', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);

    await expect(
      setPptxRoundTripShapeTransform(snapshot, {
        targetKey: 'slide-1-element-1',
        value: CHANGED_TRANSFORM,
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint shape transform target key does not exist',
    });
  });
});
