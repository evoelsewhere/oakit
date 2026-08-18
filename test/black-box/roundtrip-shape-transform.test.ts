import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  replacePptxRoundTripText,
  serializePptxRoundTripJson,
  setPptxRoundTripShapeTransform,
  setPptxRoundTripTextTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
  type PptxSceneTransform,
} from '../../src';

function scene(): PptxSceneDocument {
  return {
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
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'source-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'source-run', text: 'Before', type: 'run' },
                  ],
                  key: 'source-paragraph',
                },
              ],
            },
            type: 'text',
          },
          {
            authored: {
              fillColor: '#F97316',
              geometry: 'ellipse',
              lineColor: '#0F172A',
              lineWidth: 2,
              transform: { height: 120, width: 180, x: 420, y: 220 },
            },
            key: 'source-shape',
            name: 'Native ellipse',
            resolved: { hidden: false },
            type: 'shape',
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
  flipVertical: false,
  height: 140,
  rotation: 25,
  width: 210,
  x: 380,
  y: 190,
};

async function fixture() {
  const created = await createPptx(scene());
  return {
    bytes: created.data,
    snapshot: await readPptxRoundTrip(created.data),
  };
}

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint shape transform editing', () => {
  it('patches one slide-owned native shape and verifies the rendered output', async () => {
    const { bytes, snapshot } = await fixture();
    expect(snapshot.document.slides[0]?.elements).toMatchObject([
      { key: 'slide-1-element-1', type: 'text' },
      { key: 'slide-1-element-2', type: 'shape' },
    ]);

    const edited = await setPptxRoundTripShapeTransform(snapshot, {
      targetKey: 'slide-1-element-2',
      value: CHANGED_TRANSFORM,
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts, verified, rendered] = await Promise.all([
      payloads(bytes),
      payloads(output.data),
      readPptxRoundTrip(output.data),
      renderPptxToSvg(output.data, { slideNumbers: [1] }),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      operations: [
        { id: 'set-transform-1', kind: 'set-transform', status: 'verified' },
      ],
      patchedPartCount: 1,
      supportProfile: {
        effectiveLevel: 'R2',
        id: 'pptx-roundtrip-native-v1',
      },
    });
    expect(verified.document.slides[0]?.elements).toMatchObject([
      { key: 'slide-1-element-1', type: 'text' },
      {
        key: 'slide-1-element-2',
        resolved: { transform: CHANGED_TRANSFORM },
        type: 'shape',
      },
    ]);
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') {
        expect(result).not.toEqual(source);
      } else {
        expect(result, name).toEqual(source);
      }
    }
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('#F97316');
    expect(rendered.slides[0]?.warnings.map(({ code }) => code)).toEqual([
      'font-substitution',
    ]);
    expect(rendered.slides[0]?.warnings).not.toContainEqual(
      expect.objectContaining({ code: 'approximate-shape' }),
    );
  });

  it('survives portable JSON with its native capability binding', async () => {
    const { snapshot } = await fixture();
    const edited = await setPptxRoundTripShapeTransform(snapshot, {
      targetKey: 'slide-1-element-2',
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
    expect(restored.supportProfile.id).toBe('pptx-roundtrip-native-v1');
    expect(output.report.level).toBe('R2');
  });

  it('combines text and shape operations under the native profile', async () => {
    const { snapshot } = await fixture();
    const textEdited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const edited = await setPptxRoundTripShapeTransform(textEdited, {
      targetKey: 'slide-1-element-2',
      value: CHANGED_TRANSFORM,
    });
    const output = await writePptxRoundTrip(edited);
    const verified = await readPptxRoundTrip(output.data);

    expect(edited.supportProfile.id).toBe('pptx-roundtrip-native-v1');
    expect(output.report.operations).toMatchObject([
      { kind: 'replace-text', status: 'verified' },
      { kind: 'set-transform', status: 'verified' },
    ]);
    expect(verified.document.slides[0]?.elements).toMatchObject([
      { text: { paragraphs: [{ children: [{ text: 'After' }] }] } },
      { resolved: { transform: CHANGED_TRANSFORM }, type: 'shape' },
    ]);
  });

  it('retains the native profile when text operations follow a shape edit', async () => {
    const { snapshot } = await fixture();
    const shapeEdited = await setPptxRoundTripShapeTransform(snapshot, {
      targetKey: 'slide-1-element-2',
      value: CHANGED_TRANSFORM,
    });
    const textEdited = await replacePptxRoundTripText(shapeEdited, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const edited = await setPptxRoundTripTextTransform(textEdited, {
      targetKey: 'slide-1-element-1',
      value: { height: 90, width: 320, x: 40, y: 50 },
    });

    expect(textEdited.supportProfile.id).toBe('pptx-roundtrip-native-v1');
    expect(edited.supportProfile.id).toBe('pptx-roundtrip-native-v1');
    expect(edited.operations.map(({ targetKey }) => targetKey)).toEqual([
      'slide-1-element-2',
      'slide-1-element-1-run-1',
      'slide-1-element-1',
    ]);
  });

  it('keeps text and shape transform entry points type-specific', async () => {
    const { snapshot } = await fixture();

    await expect(
      setPptxRoundTripShapeTransform(snapshot, {
        targetKey: 'slide-1-element-1',
        value: CHANGED_TRANSFORM,
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint shape transform target key does not exist',
    });
    await expect(
      setPptxRoundTripTextTransform(snapshot, {
        targetKey: 'slide-1-element-2',
        value: CHANGED_TRANSFORM,
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint transform target key does not exist',
    });
  });

  it('requires the exact native R2 support level', async () => {
    const { snapshot } = await fixture();
    const edited = await setPptxRoundTripShapeTransform(snapshot, {
      targetKey: 'slide-1-element-2',
      value: CHANGED_TRANSFORM,
    });
    edited.supportProfile.effectiveLevel = 'R0';

    await expect(writePptxRoundTrip(edited)).rejects.toMatchObject({
      code: 'invalid-snapshot',
      message: 'PowerPoint native edit snapshot support level must be R2',
    });
  });
});
