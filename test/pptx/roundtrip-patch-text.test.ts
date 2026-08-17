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

  it('rejects a text body with multiple source nodes', async () => {
    const created = await createPptx(scene(['Before', ' second']));
    const scheduled = await schedule(created.data, 'After');

    await expect(writePptxRoundTrip(scheduled)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message: 'PowerPoint text edit target must contain exactly one text node',
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
