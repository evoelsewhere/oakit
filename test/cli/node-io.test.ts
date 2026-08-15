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
    readFile: vi.fn(() => Promise.resolve(Uint8Array.from([1, 2, 3]))),
    stdin,
    writeFile: vi.fn(() => Promise.resolve()),
    writeStderr: vi.fn(),
    writeStdout: vi.fn(),
  };
}

describe('Node CLI I/O adapter', () => {
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
