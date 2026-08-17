import { Buffer } from 'node:buffer';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { replacePptxRoundTripText } from '../../src/formats/pptx/roundtrip/edit';
import { readPptxRoundTrip } from '../../src/formats/pptx/roundtrip/read';
import { writePptxRoundTrip } from '../../src/formats/pptx/roundtrip/write';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';

function scene(textRuns: readonly string[]): PptxSceneDocument {
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
                  children: textRuns.map((text, index) => ({
                    key: `source-run-${index + 1}`,
                    text,
                    type: 'run' as const,
                  })),
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

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

async function schedule(data: Uint8Array, value: string) {
  return replacePptxRoundTripText(await readPptxRoundTrip(data), {
    targetKey: 'slide-1-element-1-run-1',
    value,
  });
}

async function rewriteSlide(
  data: Uint8Array,
  update: (xml: string) => string | Uint8Array,
  part = 'ppt/slides/slide1.xml',
): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(data);
  const entry = archive.file(part);
  if (entry === null) throw new Error(`Missing slide part ${part}`);
  archive.file(part, update(await entry.async('text')));
  return archive.generateAsync({
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
    type: 'uint8array',
  });
}

function secondSlide(document: PptxSceneDocument): void {
  document.slides.push({
    elements: [
      {
        authored: {
          transform: { height: 60, width: 240, x: 40, y: 50 },
        },
        key: 'second-text',
        resolved: { hidden: false },
        text: {
          body: {},
          paragraphs: [
            {
              children: [{ key: 'second-run', text: 'Second', type: 'run' }],
              key: 'second-paragraph',
            },
          ],
        },
        type: 'text',
      },
    ],
    key: 'second-slide',
  });
}

describe('PowerPoint part-preserving text patching', () => {
  it('changes one slide payload and verifies the complete requested preview', async () => {
    const created = await createPptx(scene(['Before']));
    const scheduled = await schedule(created.data, ' After <& ');
    const output = await writePptxRoundTrip(scheduled);
    const [sourceParts, outputParts] = await Promise.all([
      payloads(created.data),
      payloads(output.data),
    ]);

    expect(output.report).toEqual({
      addedPartCount: 0,
      copiedPartCount: 10,
      diagnostics: [],
      level: 'R2',
      operations: [
        { id: 'replace-text-1', kind: 'replace-text', status: 'verified' },
      ],
      patchedPartCount: 1,
      producerEvidence: [],
      rebuiltPartCount: 0,
      removedPartCount: 0,
      supportProfile: {
        effectiveLevel: 'R2',
        id: 'pptx-roundtrip-text-v1',
        producerMatrix: [],
        version: '1',
      },
    });
    expect(outputParts.size).toBe(sourceParts.size);
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (result === undefined) continue;
      if (name === 'ppt/slides/slide1.xml') {
        expect(result).not.toEqual(source);
        expect(new TextDecoder().decode(result)).toContain(
          '<a:t xml:space="preserve"> After &lt;&amp; </a:t>',
        );
      } else {
        expect(result, name).toEqual(source);
      }
    }
    expect(scheduled.source.data).toEqual(created.data);
  });

  it('is byte deterministic across independent writes', async () => {
    const created = await createPptx(scene(['Before']));
    const scheduled = await schedule(created.data, 'After');

    const [first, second] = await Promise.all([
      writePptxRoundTrip(scheduled),
      writePptxRoundTrip(scheduled),
    ]);

    expect(second).toEqual(first);
  });

  it('resolves a second slide through presentation relationships', async () => {
    const document = scene(['First']);
    secondSlide(document);
    const created = await createPptx(document);
    const scheduled = await replacePptxRoundTripText(
      await readPptxRoundTrip(created.data),
      { targetKey: 'slide-2-element-1-run-1', value: 'Changed second' },
    );
    const output = await writePptxRoundTrip(scheduled);
    const [sourceParts, outputParts] = await Promise.all([
      payloads(created.data),
      payloads(output.data),
    ]);

    expect(output.report.patchedPartCount).toBe(1);
    expect(outputParts.get('ppt/slides/slide1.xml')).toEqual(
      sourceParts.get('ppt/slides/slide1.xml'),
    );
    expect(outputParts.get('ppt/slides/slide2.xml')).not.toEqual(
      sourceParts.get('ppt/slides/slide2.xml'),
    );
    const verified = await readPptxRoundTrip(output.data);
    expect(verified.document.slides[1]?.elements[0]).toMatchObject({
      type: 'text',
      text: {
        paragraphs: [{ children: [{ text: 'Changed second' }] }],
      },
    });
  });

  it('patches namespace aliases without relying on conventional prefixes', async () => {
    const created = await createPptx(scene(['Before']));
    const aliased = await rewriteSlide(created.data, (xml) =>
      xml
        .replaceAll('xmlns:p=', 'xmlns:pres=')
        .replaceAll('<p:', '<pres:')
        .replaceAll('</p:', '</pres:')
        .replaceAll('xmlns:a=', 'xmlns:draw=')
        .replaceAll('<a:', '<draw:')
        .replaceAll('</a:', '</draw:'),
    );
    const output = await writePptxRoundTrip(await schedule(aliased, 'After'));
    const slide = await (
      await JSZip.loadAsync(output.data)
    )
      .file('ppt/slides/slide1.xml')
      ?.async('text');

    expect(slide).toContain('<draw:t xml:space="preserve">After</draw:t>');
    expect(slide).not.toContain('<a:t');
  });

  it('rejects a text body with multiple source nodes', async () => {
    const created = await createPptx(scene(['Before', ' second']));
    const scheduled = await schedule(created.data, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message: 'PowerPoint text edit target must contain exactly one text node',
    });
  });

  it('rejects duplicate non-visual shape identifiers', async () => {
    const created = await createPptx(scene(['Before']));
    const duplicated = await rewriteSlide(created.data, (xml) => {
      const shape = /<p:sp>.*?<\/p:sp>/.exec(xml)?.[0];
      if (shape === undefined) throw new Error('Expected source shape');
      return xml.replace('</p:spTree>', `${shape}</p:spTree>`);
    });
    const scheduled = await schedule(duplicated, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message: 'PowerPoint text edit requires one unique text shape for id 2',
    });
  });

  it.each([
    ['PresentationML extension', '<p:extLst/>'],
    ['DrawingML extension', '<a:extLst/>'],
  ])('rejects target compatibility markup: %s', async (_name, markup) => {
    const created = await createPptx(scene(['Before']));
    const extended = await rewriteSlide(created.data, (xml) =>
      xml.replace('</p:sp>', `${markup}</p:sp>`),
    );
    const scheduled = await schedule(extended, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message:
        'PowerPoint text edit target contains unsupported compatibility markup',
    });
  });

  it('rejects Office escape sequences instead of changing their semantics', async () => {
    const created = await createPptx(scene(['_x0041_']));
    const scheduled = await schedule(created.data, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message:
        'PowerPoint text edit source XML does not match its preview precondition',
    });
  });

  it('rejects UTF-16 slide XML instead of rewriting its encoding', async () => {
    const created = await createPptx(scene(['Before']));
    const utf16 = await rewriteSlide(created.data, (xml) => {
      const declared = xml.replace('encoding="UTF-8"', 'encoding="UTF-16"');
      return Uint8Array.from([0xff, 0xfe, ...Buffer.from(declared, 'utf16le')]);
    });
    const scheduled = await schedule(utf16, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message: 'PowerPoint text edit requires UTF-8 slide XML',
    });
  });

  it.each(['_xmlsignatures/sig1.xml', 'ppt/vbaProject.bin'])(
    'rejects protected package feature %s',
    async (part) => {
      const created = await createPptx(scene(['Before']));
      const archive = await JSZip.loadAsync(created.data);
      archive.file(part, new Uint8Array([1, 2, 3]));
      const protectedPackage = await archive.generateAsync({
        type: 'uint8array',
      });
      const scheduled = await schedule(protectedPackage, 'After');

      await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
        code: 'unsupported-edit-operation',
        message:
          'PowerPoint text edit does not modify signed or macro-enabled packages',
      });
    },
  );
});
