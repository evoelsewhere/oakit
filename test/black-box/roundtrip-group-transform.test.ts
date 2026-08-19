import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  setPptxRoundTripGroupTransform,
  setPptxRoundTripShapeTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
  type PptxSceneGroupTransform,
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
              transform: {
                childSpace: { height: 100, width: 100, x: 10, y: 20 },
                height: 100,
                width: 200,
                x: 100,
                y: 120,
              },
            },
            elements: [
              {
                authored: {
                  fillColor: '#F97316',
                  geometry: 'rect',
                  transform: {
                    height: 40,
                    rotation: 90,
                    width: 20,
                    x: 20,
                    y: 30,
                  },
                },
                key: 'rotated-child',
                resolved: { hidden: false },
                type: 'shape',
              },
              {
                authored: {
                  transform: {
                    childSpace: { height: 30, width: 30, x: 0, y: 0 },
                    height: 30,
                    width: 30,
                    x: 60,
                    y: 20,
                  },
                },
                elements: [
                  {
                    authored: {
                      fillColor: '#22C55E',
                      transform: { height: 10, width: 10, x: 5, y: 5 },
                    },
                    key: 'deep-child',
                    resolved: { hidden: false },
                    type: 'shape',
                  },
                ],
                key: 'nested-group',
                resolved: { hidden: false },
                type: 'group',
              },
            ],
            key: 'source-group',
            resolved: { hidden: false },
            type: 'group',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

const CHANGED_TRANSFORM: PptxSceneGroupTransform = {
  childSpace: { height: 100, width: 150, x: 0, y: 0 },
  flipHorizontal: true,
  flipVertical: false,
  height: 200,
  rotation: 15,
  width: 300,
  x: 150,
  y: 160,
};

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint group transform editing', () => {
  it('patches group coordinate spaces and verifies scaled descendants', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const edited = await setPptxRoundTripGroupTransform(snapshot, {
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
      elements: [
        {
          resolved: {
            transform: {
              height: 80,
              rotation: 90,
              width: 40,
              x: 40,
              y: 60,
            },
          },
          type: 'shape',
        },
        {
          elements: [
            {
              resolved: {
                transform: { height: 20, width: 20, x: 10, y: 10 },
              },
            },
          ],
          resolved: {
            transform: {
              childSpace: { height: 30, width: 30, x: 0, y: 0 },
              height: 60,
              width: 60,
              x: 120,
              y: 40,
            },
          },
          type: 'group',
        },
      ],
      resolved: { transform: CHANGED_TRANSFORM },
      type: 'group',
    });
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') expect(result).not.toEqual(source);
      else expect(result, name).toEqual(source);
    }
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('#F97316');
    expect(svg).toContain('#22C55E');
  });

  it('survives portable JSON with native group capability binding', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const edited = await setPptxRoundTripGroupTransform(snapshot, {
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

  it('keeps group and shape transform APIs type-specific', async () => {
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
