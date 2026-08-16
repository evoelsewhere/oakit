import { describe, expect, it } from 'vitest';

import {
  PptxRenderError,
  rasterizePptxSvgResult,
  renderPptxDocumentToPng,
  renderPptxToPng,
} from '../../src/formats/pptx/node';
import type { PptxDocument } from '../../src/formats/pptx/types';
import { createMinimalPptx } from './fixture';

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

function dimensions(data: Uint8Array): { height: number; width: number } {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return { height: view.getUint32(20), width: view.getUint32(16) };
}

function document(): PptxDocument {
  return {
    size: { height: 3, width: 2 },
    slides: [
      {
        elements: [],
        fill: { type: 'color', value: '#ff0000' },
        layoutElements: [],
        note: '',
      },
    ],
    themeColors: [],
    usedFonts: [],
  };
}

describe('PowerPoint PNG rendering in Node.js', () => {
  it('rasterizes a document into independently owned PNG bytes', () => {
    const result = renderPptxDocumentToPng(document(), { scale: 2 });

    expect(result.slides).toHaveLength(1);
    const slide = result.slides[0];
    expect(slide).toMatchObject({
      format: 'png',
      height: 6,
      mimeType: 'image/png',
      slideNumber: 1,
      warnings: [],
      width: 4,
    });
    expect(Array.from(slide?.data.slice(0, 8) ?? [])).toEqual(PNG_SIGNATURE);
    expect(slide ? dimensions(slide.data) : null).toEqual({
      height: 6,
      width: 4,
    });
  });

  it('preserves slide metadata and warnings while rasterizing SVG', () => {
    const warnings = [
      {
        code: 'font-substitution' as const,
        elementId: 'text-1',
        message: 'Portable font',
        slideNumber: 3,
      },
    ];
    const result = rasterizePptxSvgResult({
      slides: [
        {
          data: new TextEncoder().encode(
            '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="3"><rect width="2" height="3" fill="#00ff00"/></svg>',
          ),
          format: 'svg',
          height: 3,
          mimeType: 'image/svg+xml',
          slideNumber: 3,
          warnings,
          width: 2,
        },
      ],
    });

    expect(result.slides[0]).toMatchObject({
      format: 'png',
      height: 3,
      mimeType: 'image/png',
      slideNumber: 3,
      warnings,
      width: 2,
    });
    expect(result.slides[0]?.data).toBeInstanceOf(Uint8Array);
    expect(dimensions(result.slides[0]?.data ?? new Uint8Array())).toEqual({
      height: 3,
      width: 2,
    });
  });

  it('wraps malformed SVG failures in a typed render error', () => {
    let caught: unknown;
    try {
      rasterizePptxSvgResult({
        slides: [
          {
            data: new TextEncoder().encode('<svg>'),
            format: 'svg',
            height: 1,
            mimeType: 'image/svg+xml',
            slideNumber: 7,
            warnings: [],
            width: 1,
          },
        ],
      });
    } catch (error) {
      caught = error;
    }

    expect(caught).toBeInstanceOf(PptxRenderError);
    if (!(caught instanceof PptxRenderError)) {
      throw new Error('Expected a typed PowerPoint render error');
    }
    expect(caught.code).toBe('rasterization-failed');
    expect(caught.message).toBe(
      'PowerPoint slide 7 could not be rasterized as PNG',
    );
    expect(caught.cause).toBeInstanceOf(Error);
  });

  it('opens and rasterizes an independently packaged PPTX', async () => {
    const result = await renderPptxToPng(await createMinimalPptx(), {
      slideNumbers: [1],
    });

    expect(result.slides[0]).toMatchObject({
      format: 'png',
      height: 405,
      mimeType: 'image/png',
      slideNumber: 1,
      width: 720,
    });
    expect(Array.from(result.slides[0]?.data.slice(0, 8) ?? [])).toEqual(
      PNG_SIGNATURE,
    );
  });

  it('preserves an explicit empty slide selection', async () => {
    await expect(
      renderPptxToPng(await createMinimalPptx(), { slideNumbers: [] }),
    ).resolves.toEqual({ slides: [] });
  });
});
