import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { type OakitCliIo, runOakitCli } from '../../src/cli/run';

const EXPECTED_HELP = `Usage: oakit [convert] <input.pptx|-> [options]
       oakit render <input.pptx|-> --output <directory> [options]
       oakit snapshot <input.pptx|-> [--output <file>]
       oakit edit-text <input.json|-> --target <run-key> --value <text> [options]
       oakit restore <input.json|-> --output <file.pptx>

Convert a PowerPoint Open XML presentation into deterministic JSON.
Render agent-readable SVG or PNG slide previews without an Office runtime.
Preserve and restore byte-exact PowerPoint packages through portable JSON.

Convert options:
  -o, --output <file>          Write JSON to a file instead of stdout
      --strict                 Reject malformed optional OOXML content
      --pretty                 Format JSON with two-space indentation
      --document-only          Omit format metadata and diagnostics
      --image-mode <mode>      Image output: none (default) or base64

Render options:
  -o, --output <directory>     Write slide files and manifest.json
      --render-format <format> png (default) or svg
      --slides <list>          One-based comma-separated slide numbers
      --scale <number>         Positive decimal output scale (default: 1)

Snapshot options:
  -o, --output <file>          Write portable JSON instead of stdout
      --pretty                 Format portable JSON with two-space indentation

Edit text options:
  -o, --output <file>          Write edited portable JSON instead of stdout
      --target <run-key>       Stable text run key from the portable document
      --value <text>           Replacement text; use --value=-5 for leading -
      --pretty                 Format portable JSON with two-space indentation

Restore options:
  -o, --output <file>          Required PowerPoint output path

PPTX input options:
      --format <pptx>          Input format; required when reading stdin
  -h, --help                   Show this help
  -v, --version                Show the installed OAKit version
`;

const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

class MemoryCliIo implements OakitCliIo {
  readonly files = new Map<string, Uint8Array | string>();
  readError?: Error;
  stderr = '';
  stdout = '';
  stdin: Uint8Array<ArrayBufferLike> = new Uint8Array();
  writeError?: Error;

  createDirectory(): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    return Promise.resolve();
  }

  readFile(filename: string): Promise<Uint8Array> {
    if (this.readError !== undefined) return Promise.reject(this.readError);
    const value = this.files.get(filename);
    if (!(value instanceof Uint8Array)) {
      return Promise.reject(new Error(`Missing input ${filename}`));
    }
    return Promise.resolve(value);
  }

  readStdin(): Promise<Uint8Array> {
    return Promise.resolve(this.stdin);
  }

  writeStderr(value: string): void {
    this.stderr += value;
  }

  writeBinaryFile(filename: string, value: Uint8Array): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    this.files.set(filename, Uint8Array.from(value));
    return Promise.resolve();
  }

  writeStdout(value: string): void {
    this.stdout += value;
  }

  writeFile(filename: string, value: string): Promise<void> {
    if (this.writeError !== undefined) return Promise.reject(this.writeError);
    this.files.set(filename, value);
    return Promise.resolve();
  }
}

async function createCliPptx(presentationXml?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="${CONTENT_TYPES_NAMESPACE}">
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`,
  );
  zip.file(
    'ppt/presentation.xml',
    presentationXml ??
      `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}">
        <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="5143500"/>
      </p:presentation>`,
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">
      <Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/slide1.xml"/>
    </Relationships>`,
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}">
      <p:cSld><p:spTree/></p:cSld>
    </p:sld>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

function parsedJson(value: string): unknown {
  return JSON.parse(value) as unknown;
}

describe('oakit CLI contract', () => {
  it('converts a PPTX path to deterministic JSON with diagnostics', async () => {
    const io = new MemoryCliIo();
    io.files.set('deck.pptx', await createCliPptx());

    const exitCode = await runOakitCli(['deck.pptx'], io, '1.2.3');

    expect(exitCode).toBe(0);
    expect(parsedJson(io.stdout)).toMatchObject({
      diagnostics: [],
      document: {
        size: { height: 405, width: 720 },
        slides: [{}],
      },
      format: 'pptx',
    });
    expect(io.stdout).not.toContain('\n  "format"');
    expect(io.stderr).toBe('');
  });

  it('supports the explicit convert command and writes pretty JSON to a file', async () => {
    const io = new MemoryCliIo();
    io.files.set('deck.pptx', await createCliPptx());

    const exitCode = await runOakitCli(
      ['convert', 'deck.pptx', '--output', 'deck.json', '--pretty'],
      io,
      '1.2.3',
    );

    expect(exitCode).toBe(0);
    expect(io.stdout).toBe('');
    const output = io.files.get('deck.json');
    expect(output).toBeTypeOf('string');
    expect(output).toContain('\n  "format": "pptx"');
  });

  it('reads stdin only when its format is explicit', async () => {
    const missingFormatIo = new MemoryCliIo();
    missingFormatIo.stdin = await createCliPptx();

    await expect(runOakitCli(['-'], missingFormatIo, '1.2.3')).resolves.toBe(2);
    expect(parsedJson(missingFormatIo.stderr)).toMatchObject({
      error: {
        code: 'format-required',
        message: 'Reading stdin requires --format pptx',
      },
    });

    const io = new MemoryCliIo();
    io.stdin = await createCliPptx();
    await expect(
      runOakitCli(['-', '--format', 'pptx', '--document-only'], io, '1.2.3'),
    ).resolves.toBe(0);
    expect(parsedJson(io.stdout)).toMatchObject({
      size: { height: 405, width: 720 },
      slides: [{}],
    });
  });

  it('rejects unsupported formats and unknown options as usage errors', async () => {
    const unsupportedIo = new MemoryCliIo();
    await expect(
      runOakitCli(['document.docx'], unsupportedIo, '1.2.3'),
    ).resolves.toBe(2);
    expect(parsedJson(unsupportedIo.stderr)).toMatchObject({
      error: {
        code: 'unsupported-format',
        message: 'Unsupported Office format: docx',
      },
    });

    const optionIo = new MemoryCliIo();
    await expect(
      runOakitCli(['deck.pptx', '--unknown'], optionIo, '1.2.3'),
    ).resolves.toBe(2);
    expect(parsedJson(optionIo.stderr)).toMatchObject({
      error: { code: 'unknown-option', message: 'Unknown option: --unknown' },
    });
  });

  it('validates input count and option values before reading a document', async () => {
    const cases = [
      {
        args: [],
        code: 'input-required',
        message: 'A PPTX input path is required',
      },
      {
        args: ['a.pptx', 'b.pptx'],
        code: 'too-many-inputs',
        message: 'Only one input document can be converted at a time',
      },
      {
        args: ['deck.pptx', '--output'],
        code: 'missing-option-value',
        message: 'Option --output requires a value',
      },
      {
        args: ['deck.pptx', '--output', '--pretty'],
        code: 'missing-option-value',
        message: 'Option --output requires a value',
      },
      {
        args: ['deck.pptx', '--image-mode', 'blob'],
        code: 'invalid-image-mode',
        message: 'Unsupported image mode: blob',
      },
    ] as const;

    for (const testCase of cases) {
      const io = new MemoryCliIo();
      await expect(runOakitCli(testCase.args, io, '1.2.3')).resolves.toBe(2);
      expect(parsedJson(io.stderr)).toMatchObject({
        error: { code: testCase.code, message: testCase.message },
      });
    }
  });

  it('accepts both documented image modes', async () => {
    for (const mode of ['none', 'base64'] as const) {
      const io = new MemoryCliIo();
      io.files.set('deck.pptx', await createCliPptx());
      await expect(
        runOakitCli(['deck.pptx', '--image-mode', mode], io, '1.2.3'),
      ).resolves.toBe(0);
      expect(io.stderr).toBe('');
    }
  });

  it('never overwrites the input document with JSON output', async () => {
    const io = new MemoryCliIo();
    const input = await createCliPptx();
    io.files.set('deck.pptx', input);

    await expect(
      runOakitCli(['deck.pptx', '-o', 'deck.pptx'], io, '1.2.3'),
    ).resolves.toBe(2);
    expect(io.files.get('deck.pptx')).toBe(input);
    expect(parsedJson(io.stderr)).toMatchObject({
      error: {
        code: 'output-overwrites-input',
        message: 'The JSON output path must not overwrite the input document',
      },
    });
  });

  it('treats dash output as stdout rather than a filesystem path', async () => {
    const io = new MemoryCliIo();
    io.files.set('./-', await createCliPptx());

    await expect(
      runOakitCli(['./-', '--format', 'pptx', '--output', '-'], io, '1.2.3'),
    ).resolves.toBe(0);
    expect(parsedJson(io.stdout)).toMatchObject({ format: 'pptx' });
    expect(io.files.get('-')).toBeUndefined();
  });

  it('keeps stdin distinct from an output file whose name is dash', async () => {
    const io = new MemoryCliIo();
    io.stdin = await createCliPptx();
    const dashFile = `${process.cwd()}/-`;

    await expect(
      runOakitCli(['-', '--format', 'pptx', '--output', dashFile], io, '1.2.3'),
    ).resolves.toBe(0);
    expect(io.files.get(dashFile)).toBeTypeOf('string');
    expect(io.stdout).toBe('');
  });

  it('returns typed strict parse diagnostics without a stack trace', async () => {
    const io = new MemoryCliIo();
    io.files.set(
      'broken.pptx',
      await createCliPptx('<p:presentation><p:child></p:presentation>'),
    );

    await expect(
      runOakitCli(['broken.pptx', '--strict'], io, '1.2.3'),
    ).resolves.toBe(1);
    expect(parsedJson(io.stderr)).toMatchObject({
      error: {
        code: 'xml-parse-failed',
        diagnostic: { part: 'ppt/presentation.xml' },
      },
    });
    expect(io.stderr).not.toContain('    at ');
  });

  it('defaults to tolerant parsing instead of silently enabling strict mode', async () => {
    const io = new MemoryCliIo();
    io.files.set(
      'broken.pptx',
      await createCliPptx('<p:presentation><p:child></p:presentation>'),
    );

    await expect(runOakitCli(['broken.pptx'], io, '1.2.3')).resolves.toBe(0);
    expect(parsedJson(io.stdout)).toMatchObject({
      diagnostics: [
        { code: 'xml-parse-failed' },
        { code: 'invalid-document-structure' },
      ],
      format: 'pptx',
    });
    expect(io.stderr).toBe('');
  });

  it('reports unexpected parser failures as conversion errors', async () => {
    const io = new MemoryCliIo();
    io.files.set('deck.pptx', await createCliPptx());
    const failParse: typeof import('../../src/formats/pptx').parsePptxWithDiagnostics =
      () => Promise.reject(new Error('unexpected parser failure'));

    await expect(
      runOakitCli(['deck.pptx'], io, '1.2.3', { parsePptx: failParse }),
    ).resolves.toBe(1);
    expect(parsedJson(io.stderr)).toEqual({
      error: {
        code: 'conversion-failed',
        message: 'unexpected parser failure',
      },
    });
  });

  it('reports input and output I/O failures separately', async () => {
    const readIo = new MemoryCliIo();
    readIo.readError = new Error('permission denied');
    await expect(runOakitCli(['deck.pptx'], readIo, '1.2.3')).resolves.toBe(1);
    expect(parsedJson(readIo.stderr)).toMatchObject({
      error: { code: 'input-read-failed', message: 'permission denied' },
    });

    const writeIo = new MemoryCliIo();
    writeIo.files.set('deck.pptx', await createCliPptx());
    writeIo.writeError = new Error('disk full');
    await expect(
      runOakitCli(['deck.pptx', '--output', 'deck.json'], writeIo, '1.2.3'),
    ).resolves.toBe(1);
    expect(parsedJson(writeIo.stderr)).toMatchObject({
      error: { code: 'output-write-failed', message: 'disk full' },
    });
  });

  it('prints stable help and version output', async () => {
    const helpIo = new MemoryCliIo();
    await expect(runOakitCli(['--help'], helpIo, '1.2.3')).resolves.toBe(0);
    expect(helpIo.stdout).toBe(EXPECTED_HELP);
    expect(helpIo.stderr).toBe('');

    const versionIo = new MemoryCliIo();
    await expect(runOakitCli(['--version'], versionIo, '1.2.3')).resolves.toBe(
      0,
    );
    expect(versionIo.stdout).toBe('oakit 1.2.3\n');

    const shortHelpIo = new MemoryCliIo();
    await expect(runOakitCli(['-h'], shortHelpIo, '1.2.3')).resolves.toBe(0);
    expect(shortHelpIo.stdout).toBe(EXPECTED_HELP);

    const shortVersionIo = new MemoryCliIo();
    await expect(runOakitCli(['-v'], shortVersionIo, '1.2.3')).resolves.toBe(0);
    expect(shortVersionIo.stdout).toBe('oakit 1.2.3\n');
  });
});
