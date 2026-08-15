import { describe, expect, it } from 'vitest';

import { parsePptx, PptxParseError } from '../../src';
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
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        message: expect.stringContaining('maxEntries'),
      },
    });
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
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        message: expect.stringContaining('maxPartBytes'),
      },
    });
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
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        message: expect.stringContaining('maxXmlDepth'),
      },
    });
  });
});
