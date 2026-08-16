import { createHash } from 'node:crypto';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  parsePptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
} from '../../src';
import { createIndependentPptx, independentTextSlide } from './pptx-package';

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

async function partDigests(bytes: Uint8Array): Promise<Record<string, string>> {
  const archive = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const files = Object.values(archive.files)
    .filter((file) => !file.dir)
    .sort((left, right) => left.name.localeCompare(right.name));
  const digests: Record<string, string> = {};
  for (const file of files) {
    digests[file.name] = sha256(await file.async('uint8array'));
  }
  return digests;
}

describe('PowerPoint portable agent hand-off through the public API', () => {
  it('preserves an independently packaged presentation byte for byte through JSON', async () => {
    const source = await createIndependentPptx(
      {
        'customXml/agent-evidence.xml':
          '<?xml version="1.0" encoding="UTF-8"?><agent:state xmlns:agent="urn:oakit:test">opaque &amp; preserved</agent:state>',
        'ppt/slides/slide1.xml': independentTextSlide(
          'Independent portable hand-off',
        ),
      },
      { compression: 'STORE' },
    );
    const sourceDigest = sha256(source);
    const sourceParts = await partDigests(source);

    const runtime = await readPptxRoundTrip(source);
    const portable = await serializePptxRoundTripJson(runtime);
    const wireJson = JSON.stringify(portable);
    const wireValue: unknown = JSON.parse(wireJson);
    const restored = await parsePptxRoundTripJson(wireValue);
    const output = await writePptxRoundTrip(restored);

    expect(portable.source.packageBase64).toBe(
      Buffer.from(source).toString('base64'),
    );
    expect(wireJson).not.toContain('Uint8Array');
    expect(wireJson).not.toContain('ArrayBuffer');
    expect(wireJson).not.toContain('Blob');
    expect(output.data).not.toBe(source);
    expect(output.data).toEqual(source);
    expect(output.data.byteLength).toBe(source.byteLength);
    expect(sha256(output.data)).toBe(sourceDigest);
    expect(await partDigests(output.data)).toEqual(sourceParts);
    expect(output.report).toMatchObject({
      addedPartCount: 0,
      diagnostics: [],
      level: 'R0',
      operations: [],
      patchedPartCount: 0,
      rebuiltPartCount: 0,
      removedPartCount: 0,
    });
    const reparsed = await parsePptx(output.data, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    });
    expect(reparsed.size).toEqual({ height: 405, width: 720 });
    expect(reparsed.slides).toHaveLength(1);
    expect(Array.isArray(reparsed.slides[0]?.elements)).toBe(true);
  });
});
