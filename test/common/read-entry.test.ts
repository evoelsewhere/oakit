import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  readZipEntryBytes,
  type StreamableZipObject,
  type ZipEntryStream,
  ZipExpansionBudgetLimitError,
} from '../../src/common/archive/read-entry';

type DataListener = (chunk: Uint8Array) => void;
type ErrorListener = (error: unknown) => void;
type EndListener = () => void;

type StreamStep =
  | { chunk: Uint8Array; type: 'data' }
  | { error: unknown; type: 'error' }
  | { type: 'end' };

class ScriptedZipStream implements ZipEntryStream {
  readonly steps: StreamStep[];
  pauseCount = 0;
  resumeCount = 0;
  private readonly dataListeners: DataListener[] = [];
  private readonly endListeners: EndListener[] = [];
  private readonly errorListeners: ErrorListener[] = [];

  constructor(steps: StreamStep[]) {
    this.steps = steps;
  }

  on(event: 'data', listener: DataListener): ZipEntryStream;
  on(event: 'error', listener: ErrorListener): ZipEntryStream;
  on(event: 'end', listener: EndListener): ZipEntryStream;
  on(
    event: 'data' | 'end' | 'error',
    listener: DataListener | EndListener | ErrorListener,
  ): ZipEntryStream {
    if (event === 'data') this.dataListeners.push(listener);
    if (event === 'error') this.errorListeners.push(listener as ErrorListener);
    if (event === 'end') this.endListeners.push(listener as EndListener);
    return this;
  }

  pause(): ZipEntryStream {
    this.pauseCount += 1;
    return this;
  }

  resume(): ZipEntryStream {
    this.resumeCount += 1;
    for (const step of this.steps) {
      if (step.type === 'data') {
        for (const listener of this.dataListeners) listener(step.chunk);
      } else if (step.type === 'error') {
        for (const listener of this.errorListeners) listener(step.error);
      } else {
        for (const listener of this.endListeners) listener();
      }
    }
    return this;
  }
}

function scriptedFile(
  stream: ScriptedZipStream,
  uncompressedSize?: unknown,
): StreamableZipObject {
  return {
    ...(uncompressedSize === undefined ? {} : { _data: { uncompressedSize } }),
    internalStream(type) {
      expect(type).toBe('uint8array');
      return stream;
    },
    name: 'part.bin',
  };
}

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

    await expect(readZipEntryBytes(file, 5)).rejects.toMatchObject({
      actual: 11,
      limit: 5,
      message: 'Expanded ZIP entry exceeds 5 bytes',
      name: 'ZipEntrySizeLimitError',
    });
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

  it('rejects a declared oversize without starting decompression', async () => {
    const stream = new ScriptedZipStream([{ type: 'end' }]);
    const file = scriptedFile(stream, 6);

    await expect(readZipEntryBytes(file, 5)).rejects.toMatchObject({
      actual: 6,
      limit: 5,
      name: 'ZipEntrySizeLimitError',
    });
    expect(stream.resumeCount).toBe(0);
  });

  it('streams entries whose declared size is unavailable', async () => {
    const stream = new ScriptedZipStream([
      { chunk: new TextEncoder().encode('ok'), type: 'data' },
      { type: 'end' },
    ]);

    await expect(
      readZipEntryBytes(scriptedFile(stream, 'unknown'), 2),
    ).resolves.toEqual(new TextEncoder().encode('ok'));
    expect(stream.resumeCount).toBe(1);
  });

  it('assembles output correctly across multiple decompression chunks', async () => {
    const content = Uint8Array.from(
      { length: 80_000 },
      (_, index) => index % 251,
    );
    const source = new JSZip();
    source.file('part.bin', content);
    const archive = await source.generateAsync({
      compression: 'DEFLATE',
      type: 'uint8array',
    });
    const loaded = await JSZip.loadAsync(archive);

    await expect(
      readZipEntryBytes(loaded.file('part.bin')!, content.byteLength),
    ).resolves.toEqual(content);
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
    ).rejects.toMatchObject({
      actual: 11,
      limit: 5,
      message: 'Expanded ZIP data exceeds the 5 byte parse budget',
      name: 'ZipExpansionBudgetLimitError',
    });
  });

  it('never consumes another chunk after the entry limit rejects', async () => {
    const stream = new ScriptedZipStream([
      { chunk: new Uint8Array(6), type: 'data' },
      { chunk: new Uint8Array(1), type: 'data' },
      { error: new Error('late error'), type: 'error' },
      { type: 'end' },
    ]);
    const consumed: number[] = [];

    await expect(
      readZipEntryBytes(scriptedFile(stream), 5, (size) => {
        consumed.push(size);
      }),
    ).rejects.toMatchObject({ actual: 6, limit: 5 });
    expect(consumed).toEqual([]);
    expect(stream.pauseCount).toBe(1);
  });

  it('never consumes another chunk after the budget callback rejects', async () => {
    const stream = new ScriptedZipStream([
      { chunk: new Uint8Array(2), type: 'data' },
      { chunk: new Uint8Array(1), type: 'data' },
      { type: 'end' },
    ]);
    const consumed: number[] = [];

    await expect(
      readZipEntryBytes(scriptedFile(stream), 5, (size) => {
        consumed.push(size);
        throw new ZipExpansionBudgetLimitError(size, 1);
      }),
    ).rejects.toBeInstanceOf(ZipExpansionBudgetLimitError);
    expect(consumed).toEqual([2]);
    expect(stream.pauseCount).toBe(1);
  });

  it('ignores data emitted after a stream error', async () => {
    const stream = new ScriptedZipStream([
      { error: 'stream failed', type: 'error' },
      { chunk: new Uint8Array(1), type: 'data' },
      { type: 'end' },
    ]);
    const consumed: number[] = [];

    await expect(
      readZipEntryBytes(scriptedFile(stream), 5, (size) => {
        consumed.push(size);
      }),
    ).rejects.toThrow('stream failed');
    expect(consumed).toEqual([]);
  });

  it('ignores data emitted after a successful end event', async () => {
    const stream = new ScriptedZipStream([
      { type: 'end' },
      { chunk: new Uint8Array(1), type: 'data' },
    ]);
    const consumed: number[] = [];

    await expect(
      readZipEntryBytes(scriptedFile(stream), 5, (size) => {
        consumed.push(size);
      }),
    ).resolves.toEqual(new Uint8Array());
    expect(consumed).toEqual([]);
  });
});
