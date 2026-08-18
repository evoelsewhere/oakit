import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  setPptxRoundTripShapeTransform,
  setPptxRoundTripTableTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
  type PptxSceneTextBody,
  type PptxSceneTransform,
} from '../../src';

function text(key: string, value: string): PptxSceneTextBody {
  return {
    body: { anchor: 'center' },
    paragraphs: [
      {
        children: [{ key: `${key}-run`, text: value, type: 'run' }],
        key: `${key}-paragraph`,
      },
    ],
  };
}

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
              transform: { height: 100, width: 300, x: 72, y: 90 },
            },
            columns: [100, 200],
            key: 'source-table',
            resolved: { hidden: false },
            rows: [
              {
                cells: [
                  { fillColor: '#E0F2FE', text: text('a', 'Alpha') },
                  { fillColor: '#E0F2FE', text: text('b', 'Beta') },
                ],
                height: 40,
              },
              {
                cells: [
                  { fillColor: '#FFFFFF', text: text('c', 'Gamma') },
                  { fillColor: '#FFFFFF', text: text('d', 'Delta') },
                ],
                height: 60,
              },
            ],
            type: 'table',
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
  height: 150,
  rotation: 10,
  width: 400,
  x: 120,
  y: 140,
};

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint table transform editing', () => {
  it('patches only table geometry and proportionally preserves its grid', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    expect(snapshot.document.slides[0]?.elements[0]).toMatchObject({
      columns: [100, 200],
      key: 'slide-1-element-1',
      rows: [{ height: 40 }, { height: 60 }],
      type: 'table',
    });

    const edited = await setPptxRoundTripTableTransform(snapshot, {
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
      columns: [expect.closeTo(400 / 3), expect.closeTo((400 * 2) / 3)],
      resolved: { transform: CHANGED_TRANSFORM },
      rows: [{ height: 60 }, { height: 90 }],
      type: 'table',
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
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('Alpha');
    expect(svg).toContain('Delta');
    expect(svg).toContain('#E0F2FE');
  });

  it('survives portable JSON with native table capability binding', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const edited = await setPptxRoundTripTableTransform(snapshot, {
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

  it('keeps table and shape transform APIs type-specific', async () => {
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
    await expect(
      setPptxRoundTripTableTransform(snapshot, {
        targetKey: 'slide-1-element-1',
        value: { ...CHANGED_TRANSFORM, height: 0.0001, width: 0.0001 },
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message:
        'PowerPoint table transform is too small for its column and row grid',
    });
  });
});
