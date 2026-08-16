import { Buffer } from 'node:buffer';

import { describe, expect, it } from 'vitest';

import {
  type OakitCliIo,
  type OakitCliOperations,
  runOakitCli,
} from '../../src/cli/run';
import {
  PptxRoundTripPortableLimitError,
  PptxWriteError,
} from '../../src/formats/pptx';
import { createMinimalPptx } from '../pptx/fixture';

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

class PortableCliIo implements OakitCliIo {
  readonly directories = new Set<string>();
  readonly files = new Map<string, Uint8Array | string>();
  stderr = '';
  stdout = '';
  stdin: Uint8Array<ArrayBufferLike> = new Uint8Array();
  readError?: Error;
  writeError?: Error;

  createDirectory(dirname: string): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    this.directories.add(dirname);
    return Promise.resolve();
  }

  readFile(filename: string): Promise<Uint8Array> {
    if (this.readError !== undefined) return Promise.reject(this.readError);
    const value = this.files.get(filename);
    if (value instanceof Uint8Array) return Promise.resolve(value);
    if (typeof value === 'string') {
      return Promise.resolve(new TextEncoder().encode(value));
    }
    return Promise.reject(new Error(`Missing input ${filename}`));
  }

  readStdin(): Promise<Uint8Array> {
    return Promise.resolve(this.stdin);
  }

  writeBinaryFile(filename: string, value: Uint8Array): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    this.files.set(filename, Uint8Array.from(value));
    return Promise.resolve();
  }

  writeFile(filename: string, value: string): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    this.files.set(filename, value);
    return Promise.resolve();
  }

  writeStderr(value: string): void {
    this.stderr += value;
  }

  writeStdout(value: string): void {
    this.stdout += value;
  }
}

function json(value: string): Record<string, unknown> {
  return JSON.parse(value) as Record<string, unknown>;
}

describe('oakit portable PowerPoint CLI contract', () => {
  it('snapshots, restores, and renders a byte-exact PowerPoint hand-off', async () => {
    const io = new PortableCliIo();
    const source = await createMinimalPptx({
      'customXml/agent-state.xml':
        '<?xml version="1.0"?><agent xmlns="urn:oakit:test">preserve me</agent>',
    });
    io.files.set('source.pptx', source);

    await expect(
      runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const portableSource = io.files.get('handoff.json');
    expect(portableSource).toBeTypeOf('string');
    const portable = json(String(portableSource));
    expect(Object.keys(portable)).toEqual([
      'consistency',
      'document',
      'format',
      'operations',
      'schemaVersion',
      'source',
      'supportProfile',
    ]);
    expect(portableSource).not.toContain('Uint8Array');
    expect(portableSource).not.toContain('ArrayBuffer');
    expect(portableSource).not.toContain('Blob');
    expect(portable.source).toMatchObject({
      byteLength: source.byteLength,
      kind: 'base64',
      packageBase64: Buffer.from(source).toString('base64'),
    });

    await expect(
      runOakitCli(
        ['restore', 'handoff.json', '--output', 'restored.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const restored = io.files.get('restored.pptx');
    expect(restored).toBeInstanceOf(Uint8Array);
    expect(restored).toEqual(source);
    expect(restored).not.toBe(source);

    await expect(
      runOakitCli(
        ['render', 'restored.pptx', '--output', 'previews', '--scale', '0.5'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const png = io.files.get('previews/slide-1.png');
    expect(png).toBeInstanceOf(Uint8Array);
    expect(
      Array.from(png instanceof Uint8Array ? png.subarray(0, 8) : []),
    ).toEqual(PNG_SIGNATURE);
    expect(json(String(io.files.get('previews/manifest.json')))).toMatchObject({
      format: 'pptx-render',
      slides: [{ file: 'slide-1.png', slideNumber: 1 }],
      source: 'restored.pptx',
    });
    expect(io.stdout).toBe('');
    expect(io.stderr).toBe('');
  });

  it('writes pretty portable JSON to stdout from explicit-format stdin', async () => {
    const io = new PortableCliIo();
    io.stdin = await createMinimalPptx();

    await expect(
      runOakitCli(
        ['snapshot', '-', '--format', 'pptx', '--pretty'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    expect(io.stdout.startsWith('{\n  "consistency"')).toBe(true);
    expect(json(io.stdout)).toMatchObject({ format: 'pptx', schemaVersion: 1 });
    expect(io.files.size).toBe(0);
    expect(io.stderr).toBe('');
  });

  it('treats an explicit dash snapshot output as stdout', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createMinimalPptx());

    await expect(
      runOakitCli(['snapshot', 'source.pptx', '--output', '-'], io, '1.2.3'),
    ).resolves.toBe(0);
    expect(json(io.stdout)).toMatchObject({ format: 'pptx', schemaVersion: 1 });
    expect(io.files.size).toBe(1);
    expect(io.stderr).toBe('');
  });

  it('restores portable JSON from stdin without a format option', async () => {
    const sourceIo = new PortableCliIo();
    sourceIo.files.set('source.pptx', await createMinimalPptx());
    await runOakitCli(['snapshot', 'source.pptx'], sourceIo, '1.2.3');

    const restoreIo = new PortableCliIo();
    restoreIo.stdin = new TextEncoder().encode(sourceIo.stdout);
    await expect(
      runOakitCli(
        ['restore', '-', '--output', 'restored.pptx'],
        restoreIo,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    expect(restoreIo.files.get('restored.pptx')).toEqual(
      sourceIo.files.get('source.pptx'),
    );
  });

  it.each([
    {
      args: ['restore'],
      code: 'input-required',
      message: 'A portable JSON input path is required',
    },
    {
      args: ['snapshot', '-'],
      code: 'format-required',
      message: 'Reading stdin requires --format pptx',
    },
    {
      args: ['snapshot', 'source.pptx', '--strict'],
      code: 'unknown-option',
      message: 'Unknown option: --strict',
    },
    {
      args: ['snapshot', 'source.pptx', '--image-mode', 'none'],
      code: 'unknown-option',
      message: 'Unknown option: --image-mode',
    },
    {
      args: ['restore', 'handoff.json'],
      code: 'restore-output-required',
      message: 'Restoring requires a PowerPoint output file',
    },
    {
      args: ['restore', 'handoff.json', '--output', '-'],
      code: 'restore-output-required',
      message: 'Restoring requires a PowerPoint output file',
    },
    {
      args: ['restore', 'handoff.json', '--output', 'out.pptx', '--pretty'],
      code: 'unknown-option',
      message: 'Unknown option: --pretty',
    },
    {
      args: [
        'restore',
        'handoff.json',
        '--output',
        'out.pptx',
        '--format',
        'pptx',
      ],
      code: 'unknown-option',
      message: 'Unknown option: --format',
    },
    {
      args: ['snapshot', 'source.pptx', '--output', 'source.pptx'],
      code: 'output-overwrites-input',
      message: 'The JSON output path must not overwrite the input document',
    },
    {
      args: ['restore', 'handoff.json', '--output', 'handoff.json'],
      code: 'output-overwrites-input',
      message:
        'The PowerPoint output path must not overwrite the portable JSON input',
    },
  ])('rejects invalid portable usage with $code', async (testCase) => {
    const io = new PortableCliIo();

    await expect(runOakitCli(testCase.args, io, '1.2.3')).resolves.toBe(2);
    expect(json(io.stderr)).toEqual({
      error: { code: testCase.code, message: testCase.message },
    });
    expect(io.files.size).toBe(0);
  });

  it('rejects malformed JSON before portable snapshot validation', async () => {
    const io = new PortableCliIo();
    io.files.set('handoff.json', Uint8Array.from(Buffer.from('{"format":')));

    await expect(
      runOakitCli(
        ['restore', 'handoff.json', '--output', 'out.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toMatchObject({
      error: { code: 'invalid-portable-json' },
    });
    expect(io.files.has('out.pptx')).toBe(false);
  });

  it('rejects invalid UTF-8 before attempting JSON parsing', async () => {
    const io = new PortableCliIo();
    io.files.set('handoff.json', Uint8Array.from([0xc3, 0x28]));

    await expect(
      runOakitCli(
        ['restore', 'handoff.json', '--output', 'out.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'invalid-portable-json',
        message: 'Portable JSON input must be valid UTF-8',
      },
    });
    expect(io.files.has('out.pptx')).toBe(false);
  });

  it('rejects tampered portable state before writing PowerPoint bytes', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createMinimalPptx());
    await runOakitCli(
      ['snapshot', 'source.pptx', '--output', 'handoff.json'],
      io,
      '1.2.3',
    );
    const portable = json(String(io.files.get('handoff.json')));
    const document = portable.document as { size: { width: number } };
    document.size.width += 1;
    io.files.set('tampered.json', JSON.stringify(portable));

    await expect(
      runOakitCli(
        ['restore', 'tampered.json', '--output', 'out.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'snapshot-consistency-failed',
        message:
          'PowerPoint round-trip snapshot consistency does not match its bound state',
      },
    });
    expect(io.files.has('out.pptx')).toBe(false);
  });

  it('preserves strict PPTX diagnostics when snapshotting fails', async () => {
    const io = new PortableCliIo();
    io.files.set(
      'broken.pptx',
      await createMinimalPptx({
        'ppt/presentation.xml': '<p:presentation><p:child></p:presentation>',
      }),
    );

    await expect(
      runOakitCli(
        ['snapshot', 'broken.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toMatchObject({
      error: {
        code: 'xml-parse-failed',
        diagnostic: { part: 'ppt/presentation.xml' },
      },
    });
    expect(io.files.has('handoff.json')).toBe(false);
  });

  it.each(['snapshot', 'restore'] as const)(
    'reports %s input read failures before output',
    async (action) => {
      const io = new PortableCliIo();
      io.readError = new Error(`${action} input denied`);
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx', '--output', 'handoff.json']
          : ['restore', 'handoff.json', '--output', 'out.pptx'];

      await expect(runOakitCli(args, io, '1.2.3')).resolves.toBe(1);
      expect(json(io.stderr)).toEqual({
        error: {
          code: 'input-read-failed',
          message: `${action} input denied`,
        },
      });
    },
  );

  it('returns structured portable limit evidence', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createMinimalPptx());
    const serializeRoundTripJson: OakitCliOperations['serializeRoundTripJson'] =
      () =>
        Promise.reject(
          new PptxRoundTripPortableLimitError('maxDecodedBytes', 101, 100),
        );

    await expect(
      runOakitCli(['snapshot', 'source.pptx'], io, '1.2.3', {
        serializeRoundTripJson,
      }),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'portable-limit-exceeded',
        diagnostic: {
          actual: 101,
          limit: 100,
          limitName: 'maxDecodedBytes',
        },
        message:
          'PowerPoint portable snapshot limit maxDecodedBytes exceeded: 101 > 100',
      },
    });
  });

  it('returns structured portable limit evidence while restoring', async () => {
    const io = new PortableCliIo();
    io.files.set('handoff.json', '{}');
    const parseRoundTripJson: OakitCliOperations['parseRoundTripJson'] = () =>
      Promise.reject(
        new PptxRoundTripPortableLimitError('maxBase64Characters', 401, 400),
      );

    await expect(
      runOakitCli(
        ['restore', 'handoff.json', '--output', 'out.pptx'],
        io,
        '1.2.3',
        { parseRoundTripJson },
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'portable-limit-exceeded',
        diagnostic: {
          actual: 401,
          limit: 400,
          limitName: 'maxBase64Characters',
        },
        message:
          'PowerPoint portable snapshot limit maxBase64Characters exceeded: 401 > 400',
      },
    });
  });

  it('preserves typed snapshot failures from portable serialization', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createMinimalPptx());
    const serializeRoundTripJson: OakitCliOperations['serializeRoundTripJson'] =
      () =>
        Promise.reject(
          new PptxWriteError(
            'invalid-snapshot',
            'Portable snapshot failed its consistency contract',
          ),
        );

    await expect(
      runOakitCli(['snapshot', 'source.pptx'], io, '1.2.3', {
        serializeRoundTripJson,
      }),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'invalid-snapshot',
        message: 'Portable snapshot failed its consistency contract',
      },
    });
  });

  it.each(['snapshot', 'restore'] as const)(
    'reports unexpected %s failures without a stack trace',
    async (action) => {
      const io = new PortableCliIo();
      const input = await createMinimalPptx();
      io.files.set(
        action === 'snapshot' ? 'source.pptx' : 'handoff.json',
        action === 'snapshot' ? input : '{}',
      );
      const overrides: Partial<OakitCliOperations> =
        action === 'snapshot'
          ? {
              readRoundTrip: () =>
                Promise.reject(new Error('unexpected snapshot failure')),
            }
          : {
              parseRoundTripJson: () =>
                Promise.reject(new Error('unexpected restore failure')),
            };
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx']
          : ['restore', 'handoff.json', '--output', 'out.pptx'];

      await expect(runOakitCli(args, io, '1.2.3', overrides)).resolves.toBe(1);
      expect(json(io.stderr)).toEqual({
        error: {
          code: `${action}-failed`,
          message: `unexpected ${action} failure`,
        },
      });
      expect(io.stderr).not.toContain('    at ');
    },
  );

  it.each(['snapshot', 'restore'] as const)(
    'reports %s output failures without claiming success',
    async (action) => {
      const io = new PortableCliIo();
      io.files.set('source.pptx', await createMinimalPptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      io.writeError = new Error(`${action} output denied`);
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx', '--output', 'second.json']
          : ['restore', 'handoff.json', '--output', 'out.pptx'];

      await expect(runOakitCli(args, io, '1.2.3')).resolves.toBe(1);
      expect(json(io.stderr)).toEqual({
        error: {
          code: 'output-write-failed',
          message: `${action} output denied`,
        },
      });
    },
  );
});
