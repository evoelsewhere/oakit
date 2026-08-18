import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { patchPptxOperations } from '../../src/formats/pptx/roundtrip/orchestration';
import type { PptxRoundTripOperation } from '../../src/formats/pptx/roundtrip/types';
import { parse } from '../../src/formats/pptx/parser';
import { createMinimalPptx } from './fixture';

describe('PowerPoint patch orchestration', () => {
  it('rejects a parsed document with a different slide count', async () => {
    const data = await createMinimalPptx();
    const document = await parse(data, { imageMode: 'none' });
    document.slides = [];

    await expect(
      patchPptxOperations(data, document, [], resolvePptxResourceLimits()),
    ).rejects.toThrow(
      'PowerPoint text edit slide order does not match the parsed document',
    );
  });

  it('rejects unsafe operation indices and non-text document targets', async () => {
    const data = await createMinimalPptx();
    const document = await parse(data, { imageMode: 'none' });
    const slide = document.slides[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        base64: '',
        blob: '',
        borderColor: '#000000',
        borderStrokeDasharray: '0',
        borderType: 'solid',
        borderWidth: 0,
        geom: 'rect',
        height: 10,
        id: '2',
        isFlipH: false,
        isFlipV: false,
        left: 1,
        order: 0,
        ref: 'image.png',
        rotate: 0,
        top: 2,
        type: 'image',
        width: 10,
      },
    ];
    const replaceOperation: PptxRoundTripOperation = {
      expectedText: 'Before',
      id: 'replace-text-1',
      kind: 'replace-text',
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [replaceOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint text edit target is not a slide-owned text element',
    );

    const transformOperation: PptxRoundTripOperation = {
      expectedTransform: {
        flipHorizontal: false,
        flipVertical: false,
        height: 10,
        rotation: 0,
        width: 10,
        x: 1,
        y: 2,
      },
      id: 'set-transform-1',
      kind: 'set-transform',
      targetKey: 'slide-1-element-1',
      value: {
        flipHorizontal: false,
        flipVertical: false,
        height: 20,
        rotation: 0,
        width: 20,
        x: 3,
        y: 4,
      },
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [transformOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint transform target is not a slide-owned text or shape element',
    );

    transformOperation.targetKey = `slide-${'9'.repeat(20)}-element-1`;
    await expect(
      patchPptxOperations(
        data,
        document,
        [transformOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow('PowerPoint transform target index is unsafe');

    for (const malformed of [
      { ...replaceOperation, targetKey: 'bad-run-key' },
      {
        ...replaceOperation,
        targetKey: 'xslide-1-element-1-run-1',
      },
      {
        ...replaceOperation,
        targetKey: 'slide-1-element-1-run-1x',
      },
      { ...transformOperation, targetKey: 'bad-element-key' },
      { ...transformOperation, targetKey: 'xslide-1-element-1' },
      { ...transformOperation, targetKey: 'slide-1-element-1x' },
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [malformed],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow(/target is not a supported slide text/);
    }

    for (const targetKey of [
      `slide-${'9'.repeat(20)}-element-1-run-1`,
      `slide-1-element-${'9'.repeat(20)}-run-1`,
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [{ ...replaceOperation, targetKey }],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow('PowerPoint text edit target index is unsafe');
    }
  });
});
