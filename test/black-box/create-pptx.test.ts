import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  PptxWriteError,
  type PptxElement,
  type PptxFidelityLevel,
  type PptxSceneDocument,
  type PptxSceneSlide,
} from '../../src';

function textContent(element: PptxElement | undefined): string {
  if (!element || !('content' in element)) {
    throw new Error('Expected a generated text-bearing element');
  }
  return element.content;
}

function slide(key: string, text: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 80, width: 300, x: 20, y: 30 },
        },
        key: `${key}-text`,
        name: `${key} text`,
        resolved: { hidden: false },
        text: {
          body: { anchor: 'center', wrap: true },
          paragraphs: [
            {
              children: [
                {
                  key: `${key}-run`,
                  properties: { bold: true, fontSize: 18 },
                  text,
                  type: 'run',
                },
              ],
              key: `${key}-paragraph`,
            },
          ],
        },
        type: 'text',
      },
    ],
    key,
    name: `${key} name`,
  };
}

function creationScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [slide('first', 'Hello <& world'), slide('second', 'Slide two')],
    themes: [],
  };
}

describe('PowerPoint creation through the public API', () => {
  it('exposes producer-verified creation as a distinct fidelity level', () => {
    const level: PptxFidelityLevel = 'C3';

    expect(level).toBe('C3');
  });

  it('creates a strict-readable and Office-free rendered text presentation with an explicit C2 report', async () => {
    const result = await createPptx(creationScene());
    const parsed = await parsePptx(result.data, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    });

    expect(Array.from(result.data.slice(0, 4))).toEqual([
      0x50, 0x4b, 0x03, 0x04,
    ]);
    expect(parsed.size).toEqual({ height: 540, width: 960 });
    expect(parsed.slides).toHaveLength(2);
    expect(parsed.slides.map((value) => value.elements.length)).toEqual([1, 1]);
    expect(textContent(parsed.slides[0]?.elements[0])).toContain(
      'Hello&nbsp;&lt;&amp;&nbsp;world',
    );
    expect(textContent(parsed.slides[1]?.elements[0])).toContain(
      'Slide&nbsp;two',
    );
    expect(result.report).toEqual({
      addedPartCount: 13,
      copiedPartCount: 0,
      diagnostics: [],
      level: 'C2',
      operations: [],
      patchedPartCount: 0,
      producerEvidence: [],
      rebuiltPartCount: 0,
      removedPartCount: 0,
      supportProfile: {
        effectiveLevel: 'C2',
        id: 'pptx-create-text-v1',
        producerMatrix: [],
        version: '1',
      },
    });
  });

  it('is byte deterministic across concurrent public calls', async () => {
    const input = creationScene();
    const [first, second, third] = await Promise.all([
      createPptx(input),
      createPptx(input),
      createPptx(input),
    ]);

    expect(second.data).toEqual(first.data);
    expect(third.data).toEqual(first.data);
    expect(second.report).toEqual(first.report);
    expect(third.report).toEqual(first.report);
  });

  it('contains only the declared deterministic package inventory', async () => {
    const result = await createPptx(creationScene());
    const archive = await JSZip.loadAsync(result.data);

    expect(Object.keys(archive.files)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      'ppt/theme/theme1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
      'ppt/slides/slide2.xml',
      'ppt/slides/_rels/slide2.xml.rels',
    ]);
  });

  it('rejects invalid input before returning package bytes', async () => {
    const input = creationScene();
    input.slides[0]?.elements.splice(0, 1);
    const invalid = input as unknown as Record<string, unknown>;
    invalid.schemaVersion = 1;

    const promise = createPptx(input);
    await expect(promise).rejects.toBeInstanceOf(PptxWriteError);
    await expect(promise).rejects.toMatchObject({
      code: 'invalid-scene',
      issues: [
        {
          code: 'unsupported-schema-version',
          message: 'Only PowerPoint scene schema version 2 is supported',
          path: '$.schemaVersion',
        },
      ],
      message: 'PowerPoint scene is not valid for creation',
    });
    await expect(promise).rejects.not.toHaveProperty('data');
  });
});
