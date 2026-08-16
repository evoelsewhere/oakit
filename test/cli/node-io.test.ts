import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, expect, it, vi } from 'vitest';

import {
  createNodeCliIo,
  type NodeCliDependencies,
} from '../../src/cli/node-io';

async function* chunks(values: readonly unknown[]): AsyncIterable<unknown> {
  await Promise.resolve();
  yield* values;
}

function createDependencies(
  stdin: AsyncIterable<unknown> = chunks([]),
): NodeCliDependencies {
  return {
    createDirectory: vi.fn(() => Promise.resolve()),
    readFile: vi.fn(() => Promise.resolve(Uint8Array.from([1, 2, 3]))),
    stdin,
    writeBinaryFile: vi.fn(() => Promise.resolve()),
    writeFile: vi.fn(() => Promise.resolve()),
    writeStderr: vi.fn(),
    writeStdout: vi.fn(),
  };
}

describe('Node CLI I/O adapter', () => {
  it('creates nested directories and writes binary bytes through the real Node adapter', async () => {
    const temporaryDirectory = await mkdtemp(join(tmpdir(), 'oakit-node-io-'));
    try {
      const outputDirectory = join(temporaryDirectory, 'nested', 'previews');
      const outputFile = join(outputDirectory, 'slide-1.png');
      const bytes = Uint8Array.from([137, 80, 78, 71]);
      const io = createNodeCliIo();

      await io.createDirectory(outputDirectory);
      await io.writeBinaryFile(outputFile, bytes);

      expect(Array.from(await readFile(outputFile))).toEqual(Array.from(bytes));
    } finally {
      await rm(temporaryDirectory, { force: true, recursive: true });
    }
  });

  it('creates output directories through the injected filesystem', async () => {
    const dependencies = createDependencies();
    const io = createNodeCliIo(dependencies);

    await expect(io.createDirectory('previews')).resolves.toBeUndefined();
    expect(dependencies.createDirectory).toHaveBeenCalledExactlyOnceWith(
      'previews',
    );
  });

  it('forwards file reads and returns their bytes', async () => {
    const dependencies = createDependencies();
    const io = createNodeCliIo(dependencies);

    await expect(io.readFile('deck.pptx')).resolves.toEqual(
      Uint8Array.from([1, 2, 3]),
    );
    expect(dependencies.readFile).toHaveBeenCalledExactlyOnceWith('deck.pptx');
  });

  it('concatenates binary and text stdin chunks in emission order', async () => {
    const dependencies = createDependencies(
      chunks([Uint8Array.from([0, 1]), 'AB', Uint8Array.from([255])]),
    );

    await expect(createNodeCliIo(dependencies).readStdin()).resolves.toEqual(
      Uint8Array.from([0, 1, 65, 66, 255]),
    );
  });

  it('defines empty stdin as an empty byte array', async () => {
    const dependencies = createDependencies();

    await expect(createNodeCliIo(dependencies).readStdin()).resolves.toEqual(
      new Uint8Array(),
    );
  });

  it.each([null, undefined, 7, {}, new ArrayBuffer(1)])(
    'rejects unsupported stdin chunk %j',
    async (chunk) => {
      const dependencies = createDependencies(chunks([chunk]));

      await expect(createNodeCliIo(dependencies).readStdin()).rejects.toThrow(
        new TypeError('stdin emitted an unsupported data type'),
      );
    },
  );

  it('writes UTF-8 output files through the injected filesystem', async () => {
    const dependencies = createDependencies();
    const io = createNodeCliIo(dependencies);

    await expect(io.writeFile('deck.json', '{"ok":true}\n')).resolves.toBe(
      undefined,
    );
    expect(dependencies.writeFile).toHaveBeenCalledExactlyOnceWith(
      'deck.json',
      '{"ok":true}\n',
      'utf8',
    );
  });

  it('writes binary output without text encoding', async () => {
    const dependencies = createDependencies();
    const io = createNodeCliIo(dependencies);
    const bytes = Uint8Array.from([137, 80, 78, 71]);

    await expect(io.writeBinaryFile('slide-1.png', bytes)).resolves.toBe(
      undefined,
    );
    expect(dependencies.writeBinaryFile).toHaveBeenCalledExactlyOnceWith(
      'slide-1.png',
      bytes,
    );
  });

  it('forwards stdout and stderr without modifying their contents', () => {
    const dependencies = createDependencies();
    const io = createNodeCliIo(dependencies);

    io.writeStdout('result\n');
    io.writeStderr('failure\n');

    expect(dependencies.writeStdout).toHaveBeenCalledExactlyOnceWith(
      'result\n',
    );
    expect(dependencies.writeStderr).toHaveBeenCalledExactlyOnceWith(
      'failure\n',
    );
  });
});
