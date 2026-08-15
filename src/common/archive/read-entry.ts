import type JSZip from 'jszip';

export interface ZipEntryStream {
  on(event: 'data', listener: (chunk: Uint8Array) => void): ZipEntryStream;
  on(event: 'error', listener: (error: unknown) => void): ZipEntryStream;
  on(event: 'end', listener: () => void): ZipEntryStream;
  pause(): ZipEntryStream;
  resume(): ZipEntryStream;
}

export interface StreamableZipObject {
  name: string;
  _data?: { uncompressedSize?: unknown };
  internalStream(type: 'uint8array'): ZipEntryStream;
}

function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(String(value));
}

export class ZipEntrySizeLimitError extends Error {
  readonly actual: number;
  readonly limit: number;

  constructor(actual: number, limit: number) {
    super(`Expanded ZIP entry exceeds ${limit} bytes`);
    this.name = 'ZipEntrySizeLimitError';
    this.actual = actual;
    this.limit = limit;
  }
}

export class ZipExpansionBudgetLimitError extends Error {
  readonly actual: number;
  readonly limit: number;

  constructor(actual: number, limit: number) {
    super(`Expanded ZIP data exceeds the ${limit} byte parse budget`);
    this.name = 'ZipExpansionBudgetLimitError';
    this.actual = actual;
    this.limit = limit;
  }
}

/** Read an entry while stopping decompression as soon as its byte limit is exceeded. */
export function readZipEntryBytes(
  file: JSZip.JSZipObject | StreamableZipObject,
  maxBytes: number,
  consumeBytes?: (byteLength: number) => void,
): Promise<Uint8Array> {
  const expectedSize = Number(
    (file as StreamableZipObject)._data?.uncompressedSize,
  );
  if (Number.isSafeInteger(expectedSize) && expectedSize > maxBytes) {
    return Promise.reject(new ZipEntrySizeLimitError(expectedSize, maxBytes));
  }

  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let byteLength = 0;
    let settled = false;
    const stream = (file as StreamableZipObject).internalStream('uint8array');

    stream
      .on('data', (chunk) => {
        if (settled) return;
        byteLength += chunk.byteLength;
        if (byteLength > maxBytes) {
          settled = true;
          chunks.length = 0;
          stream.pause();
          reject(new ZipEntrySizeLimitError(byteLength, maxBytes));
          return;
        }
        try {
          consumeBytes?.(chunk.byteLength);
          chunks.push(chunk);
        } catch (error) {
          settled = true;
          chunks.length = 0;
          stream.pause();
          reject(asError(error));
        }
      })
      .on('error', (error) => {
        settled = true;
        chunks.length = 0;
        reject(asError(error));
      })
      .on('end', () => {
        settled = true;
        const output = new Uint8Array(byteLength);
        let offset = 0;
        for (const chunk of chunks) {
          output.set(chunk, offset);
          offset += chunk.byteLength;
        }
        resolve(output);
      })
      .resume();
  });
}
