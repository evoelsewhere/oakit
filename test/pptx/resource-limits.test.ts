import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { parsePptx, PptxParseError } from '../../src';
import {
  assertPptxArchiveWithinLimits,
  assertPptxInputWithinLimits,
  PptxResourceLimitError,
  resolvePptxResourceLimits,
  resourceLimitDiagnostic,
} from '../../src/formats/pptx/internal/resource-limits';
import { createMinimalPptx } from './fixture';

describe('PPTX resource limits', () => {
  it('rejects input larger than the configured maximum', async () => {
    const input = await createMinimalPptx();

    const result = parsePptx(input, {
      limits: { maxInputBytes: input.byteLength - 1 },
    });

    await expect(result).rejects.toBeInstanceOf(PptxParseError);
    await expect(result).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        severity: 'error',
      },
    });
  });

  it('rejects packages with too many entries', async () => {
    const input = await createMinimalPptx();

    await expect(
      parsePptx(input, { limits: { maxEntries: 1 } }),
    ).rejects.toThrow('maxEntries');
  });

  it('rejects a package part larger than the configured maximum', async () => {
    const input = await createMinimalPptx();

    await expect(
      parsePptx(input, {
        limits: {
          maxMediaBytes: 100,
          maxPartBytes: 100,
          maxXmlBytes: 100,
        },
      }),
    ).rejects.toThrow('maxPartBytes');
  });

  it('validates limit values before reading the package', async () => {
    const input = await createMinimalPptx();

    await expect(
      parsePptx(input, { limits: { maxEntries: 0 } }),
    ).rejects.toThrow(/positive integer/);
  });

  it('enforces XML complexity limits in tolerant mode', async () => {
    const input = await createMinimalPptx();

    await expect(
      parsePptx(input, {
        errorMode: 'tolerant',
        limits: { maxXmlDepth: 1 },
      }),
    ).rejects.toThrow('maxXmlDepth');
  });

  it('requires XML and media limits to fit inside the part limit', () => {
    expect(() =>
      resolvePptxResourceLimits({
        maxMediaBytes: 100,
        maxPartBytes: 100,
        maxXmlBytes: 101,
      }),
    ).toThrow('maxXmlBytes cannot exceed maxPartBytes');
    expect(() =>
      resolvePptxResourceLimits({
        maxMediaBytes: 101,
        maxPartBytes: 100,
        maxXmlBytes: 100,
      }),
    ).toThrow('maxMediaBytes cannot exceed maxPartBytes');
    expect(() =>
      resolvePptxResourceLimits({
        maxMediaBytes: 100,
        maxPartBytes: 100,
        maxXmlBytes: 100,
      }),
    ).not.toThrow();
  });

  it('allows input exactly at the configured compressed-byte limit', () => {
    const limits = resolvePptxResourceLimits({ maxInputBytes: 4 });
    expect(() =>
      assertPptxInputWithinLimits(new Uint8Array(4), limits),
    ).not.toThrow();
  });

  it('allows archive counts and sizes exactly at every configured limit', async () => {
    const source = new JSZip();
    source.file('one.bin', '12');
    source.file('two.bin', '345');
    const archive = await source.generateAsync({ type: 'uint8array' });
    const loaded = await JSZip.loadAsync(archive);
    const limits = resolvePptxResourceLimits({
      maxEntries: 2,
      maxMediaBytes: 3,
      maxPartBytes: 3,
      maxTotalUncompressedBytes: 5,
      maxXmlBytes: 3,
    });

    expect(() => assertPptxArchiveWithinLimits(loaded, limits)).not.toThrow();
  });

  it('reports cumulative expanded bytes and the offending part precisely', async () => {
    const source = new JSZip();
    source.file('one.bin', '123');
    source.file('two.bin', '456');
    const archive = await source.generateAsync({ type: 'uint8array' });
    const loaded = await JSZip.loadAsync(archive);
    const totalLimits = resolvePptxResourceLimits({
      maxEntries: 2,
      maxMediaBytes: 3,
      maxPartBytes: 3,
      maxTotalUncompressedBytes: 5,
      maxXmlBytes: 3,
    });
    try {
      assertPptxArchiveWithinLimits(loaded, totalLimits);
      throw new Error('Expected the cumulative limit to reject');
    } catch (caught) {
      expect(caught).toBeInstanceOf(PptxResourceLimitError);
      if (!(caught instanceof PptxResourceLimitError)) throw caught;
      expect(caught).toMatchObject({
        actual: 6,
        limit: 5,
        limitName: 'maxTotalUncompressedBytes',
        name: 'PptxResourceLimitError',
      });
    }

    const error = new PptxResourceLimitError(
      'maxPartBytes',
      6,
      5,
      'ppt/media/video.mp4',
    );
    expect(error.message).toBe(
      'PPTX resource limit maxPartBytes exceeded for ppt/media/video.mp4: 6 > 5',
    );
    expect(resourceLimitDiagnostic(error)).toEqual({
      code: 'resource-limit-exceeded',
      message: error.message,
      part: 'ppt/media/video.mp4',
      severity: 'error',
    });
  });

  it('rejects invalid declared expanded sizes before summing the archive', () => {
    const archive = {
      files: {
        'bad.bin': {
          _data: { uncompressedSize: -1 },
          dir: false,
          name: 'bad.bin',
        },
      },
    } as unknown as JSZip;

    expect(() =>
      assertPptxArchiveWithinLimits(archive, resolvePptxResourceLimits()),
    ).toThrow('Unable to validate expanded size for ZIP part bad.bin');
  });
});
