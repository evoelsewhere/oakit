import { Buffer } from 'node:buffer';

import { describe, expect, it } from 'vitest';

import {
  type OakitCliIo,
  type OakitCliOperations,
  runOakitCli,
} from '../../src/cli/run';
import {
  createPptx,
  PptxRoundTripPortableLimitError,
  PptxWriteError,
  readPptxRoundTrip,
  type PptxSceneDocument,
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

async function createEditablePptx(): Promise<Uint8Array> {
  const document: PptxSceneDocument = {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'cli-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'cli-run', text: 'Before CLI edit', type: 'run' },
                  ],
                  key: 'cli-paragraph',
                },
              ],
            },
            type: 'text',
          },
          {
            authored: {
              transform: { height: 40, width: 180, x: 500, y: 100 },
            },
            key: 'cli-decoy-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'cli-decoy-run', text: 'Decoy', type: 'run' },
                  ],
                  key: 'cli-decoy-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'cli-slide',
      },
    ],
    themes: [],
  };
  return (await createPptx(document)).data;
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

  it('schedules, restores, and verifies a text edit using only portable commands', async () => {
    const io = new PortableCliIo();
    const source = await createEditablePptx();
    io.files.set('source.pptx', source);
    await runOakitCli(
      ['snapshot', 'source.pptx', '--output', 'handoff.json'],
      io,
      '1.2.3',
    );

    await expect(
      runOakitCli(
        [
          'edit-text',
          'handoff.json',
          '--target',
          'slide-1-element-1-run-1',
          '--value',
          'After <& CLI edit',
          '--pretty',
          '--output',
          'edited.json',
        ],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const portable = json(String(io.files.get('edited.json')));
    expect(portable.operations).toEqual([
      {
        expectedText: 'Before CLI edit',
        id: 'replace-text-1',
        kind: 'replace-text',
        targetKey: 'slide-1-element-1-run-1',
        value: 'After <& CLI edit',
      },
    ]);
    expect(portable.source).toMatchObject({
      byteLength: source.byteLength,
      packageBase64: Buffer.from(source).toString('base64'),
    });

    await expect(
      runOakitCli(
        ['restore', 'edited.json', '--output', 'edited.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const output = io.files.get('edited.pptx');
    if (!(output instanceof Uint8Array))
      throw new Error('Expected PPTX output');
    const verified = await readPptxRoundTrip(output);
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      type: 'text',
      text: {
        paragraphs: [
          { children: [{ text: 'After <& CLI edit', type: 'run' }] },
        ],
      },
    });
    expect(output).not.toEqual(source);
    expect(io.stdout).toBe('');
    expect(io.stderr).toBe('');
  });

  it('accepts leading-hyphen replacement text using inline option syntax', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createEditablePptx());
    await runOakitCli(['snapshot', 'source.pptx'], io, '1.2.3');
    io.stdin = new TextEncoder().encode(io.stdout);
    io.stdout = '';

    await expect(
      runOakitCli(
        ['edit-text', '-', '--target=slide-1-element-1-run-1', '--value=-5'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    expect(json(io.stdout).operations).toMatchObject([{ value: '-5' }]);
    expect(io.stderr).toBe('');
  });

  it('partially transforms text through portable JSON and restores the geometry', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createEditablePptx());
    await runOakitCli(
      ['snapshot', 'source.pptx', '--output', 'handoff.json'],
      io,
      '1.2.3',
    );

    await expect(
      runOakitCli(
        [
          'transform-text',
          'handoff.json',
          '--target',
          'slide-1-element-1',
          '--x=-10',
          '--width',
          '400',
          '--rotation',
          '45',
          '--flip-horizontal',
          'true',
          '--output',
          'transformed.json',
          '--pretty',
        ],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const portable = json(String(io.files.get('transformed.json')));
    expect(portable.operations).toEqual([
      {
        expectedTransform: {
          flipHorizontal: false,
          flipVertical: false,
          height: 80,
          rotation: 0,
          width: 300,
          x: 20,
          y: 30,
        },
        id: 'set-transform-1',
        kind: 'set-transform',
        targetKey: 'slide-1-element-1',
        value: {
          flipHorizontal: true,
          flipVertical: false,
          height: 80,
          rotation: 45,
          width: 400,
          x: -10,
          y: 30,
        },
      },
    ]);

    await expect(
      runOakitCli(
        ['restore', 'transformed.json', '--output', 'transformed.pptx'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    const output = io.files.get('transformed.pptx');
    if (!(output instanceof Uint8Array))
      throw new Error('Expected PPTX output');
    const verified = await readPptxRoundTrip(output);
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      resolved: {
        transform: {
          flipHorizontal: true,
          flipVertical: false,
          height: 80,
          rotation: 45,
          width: 400,
          x: -10,
          y: 30,
        },
      },
    });
    expect(io.stderr).toBe('');
  });

  it.each([
    ['--x', '11', { x: 11 }],
    ['--y', '12', { y: 12 }],
    ['--width', '330', { width: 330 }],
    ['--height', '90', { height: 90 }],
    ['--rotation', '25', { rotation: 25 }],
    ['--flip-horizontal', 'true', { flipHorizontal: true }],
    ['--flip-vertical', 'true', { flipVertical: true }],
  ])(
    'accepts transform option %s independently',
    async (option, value, expected) => {
      const sourceIo = new PortableCliIo();
      sourceIo.files.set('source.pptx', await createEditablePptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        sourceIo,
        '1.2.3',
      );
      const io = new PortableCliIo();
      io.files.set('handoff.json', String(sourceIo.files.get('handoff.json')));

      await expect(
        runOakitCli(
          [
            'transform-text',
            'handoff.json',
            '--target=slide-1-element-1',
            option,
            value,
          ],
          io,
          '1.2.3',
        ),
      ).resolves.toBe(0);
      expect(json(io.stdout).operations).toMatchObject([{ value: expected }]);
      expect(io.stderr).toBe('');
    },
  );

  it('transforms portable JSON from stdin', async () => {
    const sourceIo = new PortableCliIo();
    sourceIo.files.set('source.pptx', await createEditablePptx());
    await runOakitCli(['snapshot', 'source.pptx'], sourceIo, '1.2.3');
    const io = new PortableCliIo();
    io.stdin = new TextEncoder().encode(sourceIo.stdout);

    await expect(
      runOakitCli(
        ['transform-text', '-', '--target=slide-1-element-1', '--x', '11'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);
    expect(json(io.stdout).operations).toMatchObject([{ value: { x: 11 } }]);
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
      args: ['edit-text'],
      code: 'input-required',
      message: 'A portable JSON input path is required',
    },
    {
      args: ['transform-text'],
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
      args: ['edit-text', 'handoff.json', '--format', 'pptx'],
      code: 'unknown-option',
      message: 'Unknown option: --format',
    },
    {
      args: ['transform-text', 'handoff.json', '--format', 'pptx'],
      code: 'unknown-option',
      message: 'Unknown option: --format',
    },
    {
      args: ['transform-text', 'handoff.json', '--value', 'After'],
      code: 'unknown-option',
      message: 'Unknown option: --value',
    },
    {
      args: ['transform-text', 'handoff.json', '--value=After'],
      code: 'unknown-option',
      message: 'Unknown option: --value=After',
    },
    {
      args: ['edit-text', 'handoff.json', '--x', '10'],
      code: 'unknown-option',
      message: 'Unknown option: --x',
    },
    {
      args: ['restore', 'handoff.json', '--target', 'key'],
      code: 'unknown-option',
      message: 'Unknown option: --target',
    },
    {
      args: ['restore', 'handoff.json', '--target=key'],
      code: 'unknown-option',
      message: 'Unknown option: --target=key',
    },
    {
      args: [
        'transform-text',
        'handoff.json',
        '--target',
        'slide-1-element-1',
        '--unknown',
      ],
      code: 'unknown-option',
      message: 'Unknown option: --unknown',
    },
    {
      args: ['edit-text', 'handoff.json', '--value', 'After'],
      code: 'edit-target-required',
      message: 'Editing text requires a non-empty --target run key',
    },
    {
      args: ['edit-text', 'handoff.json', '--target=', '--value', 'After'],
      code: 'edit-target-required',
      message: 'Editing text requires a non-empty --target run key',
    },
    {
      args: [
        'edit-text',
        'handoff.json',
        '--target',
        'slide-1-element-1-run-1',
      ],
      code: 'edit-value-required',
      message: 'Editing text requires --value',
    },
    {
      args: ['transform-text', 'handoff.json', '--x', '10'],
      code: 'transform-target-required',
      message: 'Transforming text requires a non-empty --target element key',
    },
    {
      args: ['transform-text', 'handoff.json', '--target=', '--x', '10'],
      code: 'transform-target-required',
      message: 'Transforming text requires a non-empty --target element key',
    },
    {
      args: ['transform-text', 'handoff.json', '--target', 'slide-1-element-1'],
      code: 'transform-value-required',
      message: 'Transforming text requires at least one transform option',
    },
    {
      args: [
        'transform-text',
        'handoff.json',
        '--target',
        'slide-1-element-1',
        '--x=NaN',
      ],
      code: 'invalid-transform-number',
      message: 'Option --x requires a finite number',
    },
    {
      args: [
        'transform-text',
        'handoff.json',
        '--target',
        'slide-1-element-1',
        '--flip-horizontal',
        'yes',
      ],
      code: 'invalid-transform-boolean',
      message: 'Option --flip-horizontal requires true or false',
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
    {
      args: [
        'edit-text',
        'handoff.json',
        '--target',
        'slide-1-element-1-run-1',
        '--value',
        'After',
        '--output',
        'handoff.json',
      ],
      code: 'output-overwrites-input',
      message:
        'The JSON output path must not overwrite the portable JSON input',
    },
    {
      args: [
        'transform-text',
        'handoff.json',
        '--target',
        'slide-1-element-1',
        '--x',
        '10',
        '--output',
        'handoff.json',
      ],
      code: 'output-overwrites-input',
      message:
        'The JSON output path must not overwrite the portable JSON input',
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

  it.each(['snapshot', 'restore', 'edit-text', 'transform-text'] as const)(
    'reports %s input read failures before output',
    async (action) => {
      const io = new PortableCliIo();
      io.readError = new Error(`${action} input denied`);
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx', '--output', 'handoff.json']
          : action === 'restore'
            ? ['restore', 'handoff.json', '--output', 'out.pptx']
            : action === 'edit-text'
              ? [
                  'edit-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1-run-1',
                  '--value',
                  'After',
                ]
              : [
                  'transform-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1',
                  '--x',
                  '10',
                ];

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

  it.each(['edit-text', 'transform-text'] as const)(
    'returns structured portable limit evidence while running %s',
    async (action) => {
      const io = new PortableCliIo();
      io.files.set('source.pptx', await createEditablePptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      const failure = new PptxRoundTripPortableLimitError(
        'maxDecodedBytes',
        101,
        100,
      );
      const args =
        action === 'edit-text'
          ? [
              'edit-text',
              'handoff.json',
              '--target',
              'slide-1-element-1-run-1',
              '--value',
              'After',
            ]
          : [
              'transform-text',
              'handoff.json',
              '--target',
              'slide-1-element-1',
              '--x',
              '10',
            ];

      await expect(
        runOakitCli(args, io, '1.2.3', {
          serializeRoundTripJson: () => Promise.reject(failure),
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
    },
  );

  it.each(['edit-text', 'transform-text'] as const)(
    'writes %s JSON to stdout for explicit dash output',
    async (action) => {
      const io = new PortableCliIo();
      io.files.set('source.pptx', await createEditablePptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      const args =
        action === 'edit-text'
          ? [
              'edit-text',
              'handoff.json',
              '--target',
              'slide-1-element-1-run-1',
              '--value',
              'After',
              '--output',
              '-',
            ]
          : [
              'transform-text',
              'handoff.json',
              '--target',
              'slide-1-element-1',
              '--x',
              '10',
              '--output',
              '-',
            ];

      io.stdout = '';
      await expect(runOakitCli(args, io, '1.2.3')).resolves.toBe(0);
      expect(json(io.stdout).operations).toHaveLength(1);
      expect(io.files.has('-')).toBe(false);
      expect(io.stderr).toBe('');
    },
  );

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

  it.each(['edit-text', 'transform-text'] as const)(
    'preserves typed %s operation failures',
    async (action) => {
      const io = new PortableCliIo();
      io.files.set('source.pptx', await createEditablePptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      const failure = new PptxWriteError(
        'invalid-edit-operation',
        `${action} rejected its target`,
      );
      const overrides: Partial<OakitCliOperations> =
        action === 'edit-text'
          ? { replaceRoundTripText: () => Promise.reject(failure) }
          : { setRoundTripTextTransform: () => Promise.reject(failure) };
      const args =
        action === 'edit-text'
          ? [
              'edit-text',
              'handoff.json',
              '--target',
              'slide-1-element-1-run-1',
              '--value',
              'After',
            ]
          : [
              'transform-text',
              'handoff.json',
              '--target',
              'slide-1-element-1',
              '--x',
              '10',
            ];

      await expect(runOakitCli(args, io, '1.2.3', overrides)).resolves.toBe(1);
      expect(json(io.stderr)).toEqual({
        error: {
          code: 'invalid-edit-operation',
          message: `${action} rejected its target`,
        },
      });
    },
  );

  it('rejects a missing transform target with a typed CLI error', async () => {
    const io = new PortableCliIo();
    io.files.set('source.pptx', await createEditablePptx());
    await runOakitCli(
      ['snapshot', 'source.pptx', '--output', 'handoff.json'],
      io,
      '1.2.3',
    );

    await expect(
      runOakitCli(
        ['transform-text', 'handoff.json', '--target', 'missing', '--x', '10'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'invalid-edit-operation',
        message: 'PowerPoint transform target has no resolved transform',
      },
    });
  });

  it.each(['snapshot', 'restore', 'edit-text', 'transform-text'] as const)(
    'reports unexpected %s failures without a stack trace',
    async (action) => {
      const io = new PortableCliIo();
      const input = await createEditablePptx();
      io.files.set('source.pptx', input);
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      io.files.set(
        action === 'snapshot' ? 'source.pptx' : 'handoff.json',
        action === 'snapshot' ? input : String(io.files.get('handoff.json')),
      );
      let overrides: Partial<OakitCliOperations>;
      if (action === 'snapshot') {
        overrides = {
          readRoundTrip: () =>
            Promise.reject(new Error('unexpected snapshot failure')),
        };
      } else if (action === 'restore') {
        overrides = {
          parseRoundTripJson: () =>
            Promise.reject(new Error('unexpected restore failure')),
        };
      } else if (action === 'edit-text') {
        overrides = {
          replaceRoundTripText: () =>
            Promise.reject(new Error('unexpected edit-text failure')),
        };
      } else {
        overrides = {
          setRoundTripTextTransform: () =>
            Promise.reject(new Error('unexpected transform-text failure')),
        };
      }
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx']
          : action === 'restore'
            ? ['restore', 'handoff.json', '--output', 'out.pptx']
            : action === 'edit-text'
              ? [
                  'edit-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1-run-1',
                  '--value',
                  'After',
                ]
              : [
                  'transform-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1',
                  '--x',
                  '10',
                ];

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

  it.each(['snapshot', 'restore', 'edit-text', 'transform-text'] as const)(
    'reports %s output failures without claiming success',
    async (action) => {
      const io = new PortableCliIo();
      io.files.set('source.pptx', await createEditablePptx());
      await runOakitCli(
        ['snapshot', 'source.pptx', '--output', 'handoff.json'],
        io,
        '1.2.3',
      );
      io.writeError = new Error(`${action} output denied`);
      const args =
        action === 'snapshot'
          ? ['snapshot', 'source.pptx', '--output', 'second.json']
          : action === 'restore'
            ? ['restore', 'handoff.json', '--output', 'out.pptx']
            : action === 'edit-text'
              ? [
                  'edit-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1-run-1',
                  '--value',
                  'After',
                  '--output',
                  'edited.json',
                ]
              : [
                  'transform-text',
                  'handoff.json',
                  '--target',
                  'slide-1-element-1',
                  '--x',
                  '10',
                  '--output',
                  'transformed.json',
                ];

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
