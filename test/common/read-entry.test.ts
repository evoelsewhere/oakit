import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  readZipEntryBytes,
  ZipExpansionBudgetLimitError,
  ZipEntrySizeLimitError,
} from '../../src/common/archive/read-entry';

describe('readZipEntryBytes', () => {
  it('reads an entry within its expanded byte limit', async () => {
    const zip = new JSZip();
    zip.file('part.bin', 'hello');
    const file = zip.file('part.bin')!;

    await expect(readZipEntryBytes(file, 5)).resolves.toEqual(
      new TextEncoder().encode('hello'),
    );
  });

  it('stops an expanded entry that crosses its byte limit', async () => {
    const zip = new JSZip();
    zip.file('part.bin', 'hello world');
    const file = zip.file('part.bin')!;

    await expect(readZipEntryBytes(file, 5)).rejects.toBeInstanceOf(
      ZipEntrySizeLimitError,
    );
  });

  it('rejects oversized declared content before reading it', async () => {
    const source = new JSZip();
    source.file('part.bin', 'a'.repeat(1_000));
    const archive = await source.generateAsync({
      compression: 'DEFLATE',
      type: 'uint8array',
    });
    const loaded = await JSZip.loadAsync(archive);

    await expect(
      readZipEntryBytes(loaded.file('part.bin')!, 10),
    ).rejects.toMatchObject({ actual: 1_000, limit: 10 });
  });

  it('stops when a shared expansion budget is exhausted', async () => {
    const zip = new JSZip();
    zip.file('part.bin', 'hello world');
    const file = zip.file('part.bin')!;
    let expandedBytes = 0;

    await expect(
      readZipEntryBytes(file, 100, (byteLength) => {
        expandedBytes += byteLength;
        if (expandedBytes > 5) {
          throw new ZipExpansionBudgetLimitError(expandedBytes, 5);
        }
      }),
    ).rejects.toBeInstanceOf(ZipExpansionBudgetLimitError);
  });
});
