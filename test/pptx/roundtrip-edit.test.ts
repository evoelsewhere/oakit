import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import {
  applyPptxRoundTripOperationsToPreview,
  normalizePptxRoundTripTransform,
  replacePptxRoundTripText,
  setPptxRoundTripTextTransform,
  validatePptxRoundTripReplaceTextRequest,
} from '../../src/formats/pptx/roundtrip/edit';
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

  it('applies a transform only to the exact text element target', async () => {
    const transformed = await setPptxRoundTripTextTransform(await snapshot(), {
      targetKey: 'slide-1-element-1',
      value: { height: 90, width: 320, x: 40, y: 50 },
    });
    transformed.document.slides[0]?.elements.push(
      {
        authored: {},
        key: 'decoy-text',
        resolved: {
          hidden: false,
          transform: { height: 10, width: 20, x: 1, y: 2 },
        },
        text: {
          body: {},
          paragraphs: [
            {
              children: [{ key: 'decoy-run', text: 'Decoy', type: 'run' }],
              key: 'decoy-paragraph',
            },
          ],
        },
        type: 'text',
      },
      {
        authored: {},
        feature: 'shape',
        key: 'slide-1-element-1',
        resolved: {
          hidden: false,
          transform: { height: 5, width: 6, x: 3, y: 4 },
        },
        type: 'unsupported',
      },
    );

    const preview = applyPptxRoundTripOperationsToPreview(transformed);

    expect(preview.slides[0]?.elements[0]?.resolved.transform).toEqual({
      flipHorizontal: false,
      flipVertical: false,
      height: 90,
      rotation: 0,
      width: 320,
      x: 40,
      y: 50,
    });
    expect(preview.slides[0]?.elements[1]?.resolved.transform).toEqual({
      height: 10,
      width: 20,
      x: 1,
      y: 2,
    });
    expect(preview.slides[0]?.elements[2]?.resolved.transform).toEqual({
      height: 5,
      width: 6,
      x: 3,
      y: 4,
    });
  });

  it('applies text only to the exact run target', async () => {
    const edited = await replacePptxRoundTripText(await snapshot(), {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const element = edited.document.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected editable text');
    const paragraph = element.text.paragraphs[0];
    if (paragraph === undefined) throw new Error('Expected paragraph');
    paragraph.children.push(
      { key: 'decoy-run', text: 'Decoy', type: 'run' },
      { key: 'slide-1-element-1-run-1', type: 'break' },
    );

    const preview = applyPptxRoundTripOperationsToPreview(edited);
    const previewElement = preview.slides[0]?.elements[0];
    if (previewElement?.type !== 'text')
      throw new Error('Expected preview text');
    expect(previewElement.text.paragraphs[0]?.children).toMatchObject([
      { key: 'slide-1-element-1-run-1', text: 'After', type: 'run' },
      { key: 'decoy-run', text: 'Decoy', type: 'run' },
      { key: 'slide-1-element-1-run-1', type: 'break' },
    ]);
  });

  it.each([
    ['replace-text', 'missing-run'],
    ['set-transform', 'missing-element'],
  ] as const)(
    'rejects a disappeared %s preview target',
    async (kind, targetKey) => {
      const value = await snapshot();
      value.operations =
        kind === 'replace-text'
          ? [
              {
                expectedText: 'Before',
                id: 'replace-text-1',
                kind,
                targetKey,
                value: 'After',
              },
            ]
          : [
              {
                expectedTransform: {
                  flipHorizontal: false,
                  flipVertical: false,
                  height: 80,
                  rotation: 0,
                  width: 300,
                  x: 20,
                  y: 30,
                },
                id: 'set-transform-1',
                kind,
                targetKey,
                value: {
                  flipHorizontal: false,
                  flipVertical: false,
                  height: 90,
                  rotation: 0,
                  width: 320,
                  x: 40,
                  y: 50,
                },
              },
            ];

      expect(() => applyPptxRoundTripOperationsToPreview(value)).toThrow(
        kind === 'replace-text'
          ? `PowerPoint text edit verification target disappeared: ${targetKey}`
          : `PowerPoint transform verification target disappeared: ${targetKey}`,
      );
    },
  );

  it.each([
    ['', 'PowerPoint transform target key must be a non-empty string'],
    [7, 'PowerPoint transform target key must be a non-empty string'],
    ['missing', 'PowerPoint transform target key does not exist'],
  ])('rejects transform target %j', async (targetKey, message) => {
    await expect(
      setPptxRoundTripTextTransform(await snapshot(), {
        targetKey: targetKey as string,
        value: { height: 90, width: 320, x: 40, y: 50 },
      }),
    ).rejects.toThrow(message);
  });

  it.each([null, [], 7])('rejects transform value %j', async (value) => {
    await expect(
      setPptxRoundTripTextTransform(await snapshot(), {
        targetKey: 'slide-1-element-1',
        value: value as never,
      }),
    ).rejects.toThrow('PowerPoint transform value must be an object');
  });

  it('rejects a transform target without resolved geometry', async () => {
    const value = await snapshot();
    const element = value.document.slides[0]?.elements[0];
    if (element === undefined) throw new Error('Expected element');
    delete element.resolved.transform;

    await expect(
      setPptxRoundTripTextTransform(value, {
        targetKey: element.key,
        value: { height: 90, width: 320, x: 40, y: 50 },
      }),
    ).rejects.toThrow('PowerPoint transform target has no resolved transform');
  });

  it('rejects non-string text request fields', async () => {
    await expect(
      replacePptxRoundTripText(await snapshot(), {
        targetKey: 7 as never,
        value: 'After',
      }),
    ).rejects.toThrow(
      'PowerPoint text edit target key must be a non-empty string',
    );
    await expect(
      replacePptxRoundTripText(await snapshot(), {
        targetKey: 'slide-1-element-1-run-1',
        value: 7 as never,
      }),
    ).rejects.toThrow('PowerPoint text edit value must be a string');
  });

  it('enforces replace-text limits at exact code-unit and UTF-8 boundaries', () => {
    expect(() =>
      validatePptxRoundTripReplaceTextRequest(
        { targetKey: 'run', value: '1234' },
        4,
      ),
    ).not.toThrow();
    for (const value of ['12345', '😀😀']) {
      expect(() =>
        validatePptxRoundTripReplaceTextRequest({ targetKey: 'run', value }, 4),
      ).toThrow('PowerPoint text edit value exceeds the XML part byte limit');
    }
  });

  it.each([
    { height: 10, width: 20, x: '1', y: 2 },
    { height: 10, width: 20, x: Number.NaN, y: 2 },
    { height: 10, rotation: '1', width: 20, x: 1, y: 2 },
    { height: 10, rotation: Number.NaN, width: 20, x: 1, y: 2 },
  ])('rejects invalid transform directly: %j', (value) => {
    expect(() => normalizePptxRoundTripTransform(value as never)).toThrow(
      'PowerPoint transform value is not a valid scene transform',
    );
  });

  it('normalizes optional transform fields directly', () => {
    expect(
      normalizePptxRoundTripTransform({
        height: 10,
        width: 20,
        x: 1,
        y: 2,
      }),
    ).toEqual({
      flipHorizontal: false,
      flipVertical: false,
      height: 10,
      rotation: 0,
      width: 20,
      x: 1,
      y: 2,
    });
  });
});
