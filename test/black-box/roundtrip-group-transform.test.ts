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
  childSpace: { height: 80, width: 150, x: 7, y: 11 },
  flipHorizontal: true,
  flipVertical: false,
  height: 320,
  rotation: 15,
  width: 450,
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
              height: 120,
              rotation: 90,
              width: 80,
              x: 29,
              y: 96,
            },
          },
          type: 'shape',
        },
        {
          elements: [
            {
              resolved: {
                transform: { height: 40, width: 30, x: 15, y: 20 },
              },
            },
          ],
          resolved: {
            transform: {
              childSpace: { height: 30, width: 30, x: 0, y: 0 },
              height: 120,
              width: 90,
              x: 159,
              y: 36,
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

  it('rescales negative right-angle rotations through both coordinate spaces', async () => {
    const negativeRotationScene = scene();
    const group = negativeRotationScene.slides[0]?.elements[0];
    if (group?.type !== 'group') throw new Error('Expected group');
    const child = group.elements[0];
    if (child === undefined) throw new Error('Expected group child');
    const childTransform = child.authored.transform;
    if (childTransform === undefined) throw new Error('Expected child transform');
    childTransform.rotation = -90;
    const snapshot = await readPptxRoundTrip(
      (await createPptx(negativeRotationScene)).data,
    );
    const edited = await setPptxRoundTripGroupTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });

    const preview = await readPptxRoundTrip(
      (await writePptxRoundTrip(edited)).data,
    );
    const previewGroup = preview.document.slides[0]?.elements[0];
    if (previewGroup?.type !== 'group') throw new Error('Expected group');
    expect(previewGroup.elements[0]).toMatchObject({
      resolved: {
        transform: {
          height: 120,
          rotation: -90,
          width: 80,
          x: 29,
          y: 96,
        },
      },
    });
  });

  it('rejects malformed group coordinate spaces in portable snapshots', async () => {
    const snapshot = await readPptxRoundTrip((await createPptx(scene())).data);
    const edited = await setPptxRoundTripGroupTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });
    const malformedChildSpaces: unknown[] = [
      null,
      { height: 80, width: 150, x: 7 },
      { extra: 1, height: 80, width: 150, x: 7, y: 11 },
      { height: 80, width: 150, x: Number.NaN, y: 11 },
      { height: 80, width: 150, x: 7, y: Number.POSITIVE_INFINITY },
      { height: 80, width: 0, x: 7, y: 11 },
      { height: 0, width: 150, x: 7, y: 11 },
    ];
    for (const field of ['value', 'expectedTransform'] as const) {
      for (const childSpace of malformedChildSpaces) {
        const invalid = structuredClone(edited);
        const operation = invalid.operations[0];
        if (operation?.kind !== 'set-transform') {
          throw new Error('Expected transform operation');
        }
        operation[field] = {
          ...(operation[field] as PptxSceneGroupTransform),
          childSpace,
        } as never;

        await expect(writePptxRoundTrip(invalid)).rejects.toMatchObject({
          code: 'invalid-snapshot',
        });
      }
    }
  });
});
