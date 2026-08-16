import { describe, expect, it } from 'vitest';

import {
  type OakitCliIo,
  type OakitCliOperations,
  runOakitCli,
} from '../../src/cli/run';
import type { PptxInputRenderOptions } from '../../src/formats/pptx';
import { createMinimalPptx } from '../pptx/fixture';

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

class RenderCliIo implements OakitCliIo {
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
    return value instanceof Uint8Array
      ? Promise.resolve(value)
      : Promise.reject(new Error(`Missing input ${filename}`));
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

function json(value: string): unknown {
  return JSON.parse(value) as unknown;
}

describe('oakit render CLI contract', () => {
  it('renders selected slides as PNG files with an agent-readable manifest', async () => {
    const io = new RenderCliIo();
    io.files.set('deck.pptx', await createMinimalPptx());

    await expect(
      runOakitCli(
        [
          'render',
          'deck.pptx',
          '--output',
          'previews',
          '--render-format',
          'png',
          '--slides',
          '1',
          '--scale',
          '0.5',
        ],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);

    expect(io.directories).toEqual(new Set(['previews']));
    const png = io.files.get('previews/slide-1.png');
    expect(png).toBeInstanceOf(Uint8Array);
    expect(
      Array.from(png instanceof Uint8Array ? png.slice(0, 8) : []),
    ).toEqual(PNG_SIGNATURE);
    const manifest = io.files.get('previews/manifest.json');
    expect(manifest).toBeTypeOf('string');
    expect(json(typeof manifest === 'string' ? manifest : '')).toMatchObject({
      format: 'pptx-render',
      renderFormat: 'png',
      scale: 0.5,
      slides: [
        {
          byteLength: png instanceof Uint8Array ? png.byteLength : -1,
          file: 'slide-1.png',
          format: 'png',
          height: 203,
          mimeType: 'image/png',
          slideNumber: 1,
          warnings: [{ code: 'font-substitution', slideNumber: 1 }],
          width: 360,
        },
      ],
      source: 'deck.pptx',
    });
    expect(io.stdout).toBe('');
    expect(io.stderr).toBe('');
  });

  it('renders stdin as a safe SVG when the input format is explicit', async () => {
    const io = new RenderCliIo();
    io.stdin = await createMinimalPptx();

    await expect(
      runOakitCli(
        [
          'render',
          '-',
          '--format',
          'pptx',
          '--output',
          'svg-output',
          '--render-format',
          'svg',
        ],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(0);

    const svg = io.files.get('svg-output/slide-1.svg');
    expect(svg).toBeInstanceOf(Uint8Array);
    const source =
      svg instanceof Uint8Array ? new TextDecoder().decode(svg) : '';
    expect(source).toContain('<title>PowerPoint slide 1</title>');
    expect(source).toContain('Hello AI');
    expect(source).not.toMatch(/<(?:script|foreignObject)\b/i);
    expect(
      json(String(io.files.get('svg-output/manifest.json'))),
    ).toMatchObject({
      renderFormat: 'svg',
      source: 'stdin',
    });
  });

  it('forwards exact multi-slide selection and supported decimal scales', async () => {
    const received: PptxInputRenderOptions[] = [];
    const renderSvg: OakitCliOperations['renderSvg'] = (
      _input,
      options = {},
    ) => {
      received.push(structuredClone(options));
      return Promise.resolve({ slides: [] });
    };
    const scales = [
      ['1', 1],
      ['12', 12],
      ['1.25', 1.25],
      ['0.25', 0.25],
    ] as const;

    for (const [source, expected] of scales) {
      const io = new RenderCliIo();
      io.files.set('deck.pptx', Uint8Array.from([1, 2, 3]));
      await expect(
        runOakitCli(
          [
            'render',
            'deck.pptx',
            '--output',
            'out',
            '--render-format',
            'svg',
            '--slides',
            '1,23',
            '--scale',
            source,
          ],
          io,
          '1.2.3',
          { renderSvg },
        ),
      ).resolves.toBe(0);
      expect(received.at(-1)).toEqual({
        scale: expected,
        slideNumbers: [1, 23],
      });
    }

    const defaultIo = new RenderCliIo();
    defaultIo.files.set('deck.pptx', Uint8Array.from([1]));
    await expect(
      runOakitCli(
        ['render', 'deck.pptx', '--output', 'out', '--render-format', 'svg'],
        defaultIo,
        '1.2.3',
        { renderSvg },
      ),
    ).resolves.toBe(0);
    const defaultOptions = received.at(-1);
    expect(defaultOptions).toEqual({ scale: 1 });
    expect(Object.hasOwn(defaultOptions ?? {}, 'slideNumbers')).toBe(false);
  });

  it.each([
    {
      args: ['render', 'deck.pptx'],
      code: 'render-output-required',
      message: 'Rendering requires an output directory',
    },
    {
      args: ['render', 'deck.pptx', '--output', '-'],
      code: 'render-output-required',
      message: 'Rendering requires an output directory',
    },
    {
      args: [
        'render',
        'deck.pptx',
        '--output',
        'out',
        '--render-format',
        'webp',
      ],
      code: 'invalid-render-format',
      message: 'Unsupported render format: webp',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', '1,0'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', '1,1'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', 'x1'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', '1x'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', ' 1'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', '1 '],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--slides', '+1'],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: [
        'render',
        'deck.pptx',
        '--output',
        'out',
        '--slides',
        '1,9007199254740992',
      ],
      code: 'invalid-slides',
      message: 'Render slides must be unique positive safe integers',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', 'Infinity'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '0'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '+1'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', 'x1'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1x'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', ' 1'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1 '],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1.'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '.25'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '.x'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1e*2'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1e'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--scale', '1ex'],
      code: 'invalid-scale',
      message: 'Render scale must be a positive finite number',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--image-mode', 'none'],
      code: 'unknown-option',
      message: 'Unknown option: --image-mode',
    },
    {
      args: ['deck.pptx', '--render-format', 'svg'],
      code: 'unknown-option',
      message: 'Unknown option: --render-format',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--document-only'],
      code: 'unknown-option',
      message: 'Unknown option: --document-only',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--strict'],
      code: 'unknown-option',
      message: 'Unknown option: --strict',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'out', '--pretty'],
      code: 'unknown-option',
      message: 'Unknown option: --pretty',
    },
    {
      args: ['deck.pptx', '--slides', '1'],
      code: 'unknown-option',
      message: 'Unknown option: --slides',
    },
    {
      args: ['deck.pptx', '--scale', '1'],
      code: 'unknown-option',
      message: 'Unknown option: --scale',
    },
    {
      args: ['render', 'deck.pptx', '--output', 'deck.pptx'],
      code: 'output-overwrites-input',
      message:
        'The render output directory must not overwrite the input document',
    },
  ])(
    'rejects invalid render usage with $code',
    async ({ args, code, message }) => {
      const io = new RenderCliIo();

      await expect(runOakitCli(args, io, '1.2.3')).resolves.toBe(2);
      expect(json(io.stderr)).toEqual({ error: { code, message } });
      expect(io.directories.size).toBe(0);
    },
  );

  it('returns a typed render error when a requested slide does not exist', async () => {
    const io = new RenderCliIo();
    io.files.set('deck.pptx', await createMinimalPptx());

    await expect(
      runOakitCli(
        ['render', 'deck.pptx', '--output', 'out', '--slides', '2'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'slide-not-found',
        message: 'PowerPoint slide 2 does not exist',
      },
    });
    expect(io.stderr).not.toContain('    at ');
    expect(io.directories.size).toBe(0);
  });

  it('reports render input reads before creating output state', async () => {
    const io = new RenderCliIo();
    io.readError = new Error('render input denied');

    await expect(
      runOakitCli(['render', 'deck.pptx', '--output', 'out'], io, '1.2.3'),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: { code: 'input-read-failed', message: 'render input denied' },
    });
    expect(io.directories.size).toBe(0);
  });

  it('preserves strict package diagnostics from the renderer', async () => {
    const io = new RenderCliIo();
    io.files.set(
      'broken.pptx',
      await createMinimalPptx({
        'ppt/presentation.xml': '<p:presentation><p:child></p:presentation>',
      }),
    );

    await expect(
      runOakitCli(['render', 'broken.pptx', '--output', 'out'], io, '1.2.3'),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toMatchObject({
      error: {
        code: 'xml-parse-failed',
        diagnostic: { part: 'ppt/presentation.xml' },
      },
    });
    expect(io.directories.size).toBe(0);
  });

  it('reports unexpected renderer failures without a stack trace', async () => {
    const io = new RenderCliIo();
    io.files.set('deck.pptx', Uint8Array.from([1]));
    const renderSvg: OakitCliOperations['renderSvg'] = () =>
      Promise.reject(new Error('unexpected renderer failure'));

    await expect(
      runOakitCli(
        ['render', 'deck.pptx', '--output', 'out', '--render-format', 'svg'],
        io,
        '1.2.3',
        { renderSvg },
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: {
        code: 'render-failed',
        message: 'unexpected renderer failure',
      },
    });
    expect(io.stderr).not.toContain('    at ');
  });

  it('reports binary output failures without claiming success', async () => {
    const io = new RenderCliIo();
    io.files.set('deck.pptx', await createMinimalPptx());
    io.writeError = new Error('read-only output');

    await expect(
      runOakitCli(
        ['render', 'deck.pptx', '--output', 'out', '--render-format', 'svg'],
        io,
        '1.2.3',
      ),
    ).resolves.toBe(1);
    expect(json(io.stderr)).toEqual({
      error: { code: 'output-write-failed', message: 'read-only output' },
    });
  });
});
