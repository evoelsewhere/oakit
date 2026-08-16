import { createHash } from 'node:crypto';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { ZipEntrySizeLimitError } from '../../src/common/archive/read-entry';
import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import {
  detectPptxRoundTripConformance,
  inspectPptxRoundTripPackage,
  normalizePptxRoundTripInput,
} from '../../src/formats/pptx/roundtrip/source';

const STRICT_NAMESPACE = 'http://purl.oclc.org/ooxml/presentationml/main';
const TRANSITIONAL_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

async function packageWithPresentation(xml?: string): Promise<Uint8Array> {
  const archive = new JSZip();
  if (xml !== undefined) archive.file('ppt/presentation.xml', xml);
  return archive.generateAsync({ type: 'uint8array' });
}

describe('PowerPoint round-trip source normalization', () => {
  it('copies only the addressed Uint8Array range before hashing', async () => {
    const owner = new Uint8Array([99, 1, 2, 3, 88]);
    const input = owner.subarray(1, 4);
    const result = await normalizePptxRoundTripInput(
      input,
      resolvePptxResourceLimits(),
    );
    owner.fill(0);

    expect(result.byteLength).toBe(3);
    expect(result.bytes).toEqual(new Uint8Array([1, 2, 3]));
    expect(result.data).toBe(result.bytes);
    expect(result.sha256).toBe(sha256(new Uint8Array([1, 2, 3])));
  });

  it('owns an independent copy of ArrayBuffer input', async () => {
    const owner = new Uint8Array([4, 5, 6]);
    const result = await normalizePptxRoundTripInput(
      owner.buffer,
      resolvePptxResourceLimits(),
    );
    owner[0] = 0;

    expect(result.bytes).toEqual(new Uint8Array([4, 5, 6]));
    expect(result.sha256).toBe(sha256(new Uint8Array([4, 5, 6])));
  });

  it('preserves Blob transport without sharing mutable byte storage', async () => {
    const input = new Blob([new Uint8Array([7, 8, 9])], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const result = await normalizePptxRoundTripInput(
      input,
      resolvePptxResourceLimits(),
    );

    expect(result.data).toBeInstanceOf(Blob);
    expect(result.data).not.toBe(input);
    if (!(result.data instanceof Blob)) throw new Error('Expected Blob data');
    expect(result.data.type).toBe(input.type);
    expect(new Uint8Array(await result.data.arrayBuffer())).toEqual(
      new Uint8Array([7, 8, 9]),
    );
    expect(result.bytes).toEqual(new Uint8Array([7, 8, 9]));
    expect(result.sha256).toBe(sha256(new Uint8Array([7, 8, 9])));
  });

  it('rejects oversized input before cloning or reading it', async () => {
    await expect(
      normalizePptxRoundTripInput(
        new Uint8Array([1, 2]),
        resolvePptxResourceLimits({ maxInputBytes: 1 }),
      ),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxInputBytes',
    });
  });
});

describe('PowerPoint round-trip conformance detection', () => {
  it.each([
    [
      'strict',
      `<?xml version="1.0"?><q:presentation xmlns:q="${STRICT_NAMESPACE}" xmlns:p="${TRANSITIONAL_NAMESPACE}"/>`,
    ],
    ['strict', `<presentation xmlns="${STRICT_NAMESPACE}"></presentation>`],
    ['transitional', `<p:presentation xmlns:p="${TRANSITIONAL_NAMESPACE}"/>`],
    ['unknown', '<p:presentation xmlns:p="urn:unknown:presentation"/>'],
    ['unknown', '<p:notPresentation xmlns:p="' + STRICT_NAMESPACE + '"/>'],
    ['unknown', '<presentation/>'],
  ] as const)(
    'detects %s from the bound root namespace',
    async (expected, xml) => {
      const bytes = await packageWithPresentation(xml);

      await expect(
        detectPptxRoundTripConformance(bytes, resolvePptxResourceLimits()),
      ).resolves.toBe(expected);
    },
  );

  it('returns unknown when the presentation part is absent', async () => {
    const bytes = await packageWithPresentation();

    await expect(
      detectPptxRoundTripConformance(bytes, resolvePptxResourceLimits()),
    ).resolves.toBe('unknown');
  });

  it('reports only non-directory package parts', async () => {
    const archive = new JSZip();
    archive.folder('empty-directory');
    archive.file(
      'ppt/presentation.xml',
      `<p:presentation xmlns:p="${STRICT_NAMESPACE}"/>`,
    );
    archive.file('docProps/core.xml', '<core/>');
    const bytes = await archive.generateAsync({ type: 'uint8array' });

    await expect(
      inspectPptxRoundTripPackage(bytes, resolvePptxResourceLimits()),
    ).resolves.toEqual({ conformance: 'strict', partCount: 2 });
  });

  it('bounds presentation XML decompression', async () => {
    const bytes = await packageWithPresentation(
      `<p:presentation xmlns:p="${TRANSITIONAL_NAMESPACE}"/>`,
    );

    await expect(
      detectPptxRoundTripConformance(
        bytes,
        resolvePptxResourceLimits({ maxXmlBytes: 10 }),
      ),
    ).rejects.toBeInstanceOf(ZipEntrySizeLimitError);
  });

  it.each([
    ['maxXmlDepth', { maxXmlDepth: 1 }],
    ['maxXmlNodes', { maxXmlNodes: 1 }],
  ] as const)(
    'enforces %s while inspecting the root',
    async (limitName, limit) => {
      const bytes = await packageWithPresentation(
        `<p:presentation xmlns:p="${TRANSITIONAL_NAMESPACE}"><p:sldIdLst/></p:presentation>`,
      );

      await expect(
        detectPptxRoundTripConformance(bytes, resolvePptxResourceLimits(limit)),
      ).rejects.toMatchObject({ limitName });
    },
  );

  it('rejects malformed XML instead of guessing its namespace', async () => {
    const bytes = await packageWithPresentation(
      `<p:presentation xmlns:p="${STRICT_NAMESPACE}">`,
    );

    await expect(
      detectPptxRoundTripConformance(bytes, resolvePptxResourceLimits()),
    ).rejects.toThrow('Invalid XML structure');
  });
});
