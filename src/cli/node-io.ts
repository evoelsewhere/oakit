import { readFile, writeFile } from 'node:fs/promises';

import type { OakitCliIo } from './run';

async function readStdin(): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  const stdin: AsyncIterable<unknown> = process.stdin;
  for await (const chunk of stdin) {
    const bytes =
      typeof chunk === 'string'
        ? new TextEncoder().encode(chunk)
        : chunk instanceof Uint8Array
          ? chunk
          : undefined;
    if (bytes === undefined) {
      throw new TypeError('stdin emitted an unsupported data type');
    }
    chunks.push(bytes);
    byteLength += bytes.byteLength;
  }

  const input = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    input.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return input;
}

export function createNodeCliIo(): OakitCliIo {
  return {
    async readFile(filename) {
      return readFile(filename);
    },
    readStdin,
    async writeFile(filename, value) {
      await writeFile(filename, value, 'utf8');
    },
    writeStderr(value) {
      process.stderr.write(value);
    },
    writeStdout(value) {
      process.stdout.write(value);
    },
  };
}
