import { mkdir, readFile, writeFile } from 'node:fs/promises';

import type { OakitCliIo } from './run';

export interface NodeCliDependencies {
  readonly createDirectory: (dirname: string) => Promise<void>;
  readonly stdin: AsyncIterable<unknown>;
  readonly readFile: (filename: string) => Promise<Uint8Array>;
  readonly writeBinaryFile: (
    filename: string,
    value: Uint8Array,
  ) => Promise<void>;
  readonly writeFile: (
    filename: string,
    value: string,
    encoding: 'utf8',
  ) => Promise<void>;
  readonly writeStderr: (value: string) => void;
  readonly writeStdout: (value: string) => void;
}

async function readStdin(stdin: AsyncIterable<unknown>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
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

const nodeCliDependencies: NodeCliDependencies = {
  async createDirectory(dirname) {
    await mkdir(dirname, { recursive: true });
  },
  readFile,
  stdin: process.stdin,
  async writeBinaryFile(filename, value) {
    await writeFile(filename, value);
  },
  writeFile,
  writeStderr: process.stderr.write.bind(process.stderr),
  writeStdout: process.stdout.write.bind(process.stdout),
};

export function createNodeCliIo(
  dependencies: NodeCliDependencies = nodeCliDependencies,
): OakitCliIo {
  return {
    async createDirectory(dirname) {
      await dependencies.createDirectory(dirname);
    },
    async readFile(filename) {
      return dependencies.readFile(filename);
    },
    async readStdin() {
      return readStdin(dependencies.stdin);
    },
    async writeBinaryFile(filename, value) {
      await dependencies.writeBinaryFile(filename, value);
    },
    async writeFile(filename, value) {
      await dependencies.writeFile(filename, value, 'utf8');
    },
    writeStderr(value) {
      dependencies.writeStderr(value);
    },
    writeStdout(value) {
      dependencies.writeStdout(value);
    },
  };
}
