import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { replacePptxRoundTripText } from '../../src/formats/pptx/roundtrip/edit';
import { readPptxRoundTrip } from '../../src/formats/pptx/roundtrip/read';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';

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
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

async function snapshot() {
  const created = await createPptx(scene());
  return readPptxRoundTrip(created.data);
}

describe('PowerPoint round-trip text edit binding', () => {
  it('binds an exact preview precondition without mutating source state', async () => {
    const source = await snapshot();
    const before = structuredClone(source);
    const edited = await replacePptxRoundTripText(source, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After <&',
    });

    expect(source).toEqual(before);
    expect(edited.document).toEqual(source.document);
    expect(edited.source).toEqual(source.source);
    expect(edited.operations).toEqual([
      {
        expectedText: 'Before',
        id: 'replace-text-1',
        kind: 'replace-text',
        targetKey: 'slide-1-element-1-run-1',
        value: 'After <&',
      },
    ]);
    expect(edited.supportProfile).toEqual({
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-text-v1',
      producerMatrix: [],
      version: '1',
    });
    expect(edited.consistency.operationsSha256).not.toBe(
      source.consistency.operationsSha256,
    );
    expect(edited.consistency.semanticPreviewSha256).toBe(
      source.consistency.semanticPreviewSha256,
    );
  });

  it('is deterministic across isolated calls', async () => {
    const source = await snapshot();
    const [first, second] = await Promise.all([
      replacePptxRoundTripText(source, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'After',
      }),
      replacePptxRoundTripText(source, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'After',
      }),
    ]);

    expect(second).toEqual(first);
  });

  it.each([
    ['', 'After', 'target key must be a non-empty string'],
    ['missing', 'After', 'target key does not exist'],
    ['slide-1-element-1-run-1', 'Before', 'must change the target value'],
    ['slide-1-element-1-run-1', 'bad\u0000', 'is not safe XML text'],
  ])(
    'rejects target %s with an invalid edit',
    async (targetKey, value, message) => {
      const editing = replacePptxRoundTripText(await snapshot(), {
        targetKey,
        value,
      });
      await expect(editing).rejects.toMatchObject({
        code: 'invalid-edit-operation',
      });
      await expect(editing).rejects.toThrow(message);
    },
  );

  it('rejects a second operation for the same target', async () => {
    const first = await replacePptxRoundTripText(await snapshot(), {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });

    await expect(
      replacePptxRoundTripText(first, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'Again',
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint text edit target is already scheduled',
    });
  });
});
