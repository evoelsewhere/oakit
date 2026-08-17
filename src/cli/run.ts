import path from 'node:path';

import {
  parsePptxRoundTripJson,
  parsePptxWithDiagnostics,
  PptxParseError,
  PptxRenderError,
  PptxRoundTripPortableLimitError,
  PptxWriteError,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  type PptxDiagnostic,
  writePptxRoundTrip,
} from '../formats/pptx';
import { renderPptxToPng } from '../formats/pptx/node';

export interface OakitCliIo {
  createDirectory(dirname: string): Promise<void>;
  readFile(filename: string): Promise<Uint8Array>;
  readStdin(): Promise<Uint8Array>;
  writeBinaryFile(filename: string, value: Uint8Array): Promise<void>;
  writeFile(filename: string, value: string): Promise<void>;
  writeStderr(value: string): void;
  writeStdout(value: string): void;
}

type OakitCliFormat = 'pptx';
type OakitRenderFormat = 'png' | 'svg';

interface ConvertCommand {
  action: 'convert';
  documentOnly: boolean;
  format: OakitCliFormat;
  imageMode: 'base64' | 'none';
  input: string;
  output?: string;
  pretty: boolean;
  strict: boolean;
}

interface RenderCommand {
  action: 'render';
  format: OakitCliFormat;
  input: string;
  output: string;
  renderFormat: OakitRenderFormat;
  scale: number;
  slideNumbers?: readonly number[];
}

interface SnapshotCommand {
  action: 'snapshot';
  format: OakitCliFormat;
  input: string;
  output?: string;
  pretty: boolean;
}

interface RestoreCommand {
  action: 'restore';
  input: string;
  output: string;
}

interface EditTextCommand {
  action: 'edit-text';
  input: string;
  output?: string;
  pretty: boolean;
  targetKey: string;
  value: string;
}

type CliCommand =
  | ConvertCommand
  | EditTextCommand
  | RenderCommand
  | RestoreCommand
  | SnapshotCommand
  | { action: 'help' }
  | { action: 'version' };

export interface OakitCliOperations {
  parseRoundTripJson: typeof parsePptxRoundTripJson;
  parsePptx: typeof parsePptxWithDiagnostics;
  readRoundTrip: typeof readPptxRoundTrip;
  replaceRoundTripText: typeof replacePptxRoundTripText;
  renderPng: typeof renderPptxToPng;
  renderSvg: typeof renderPptxToSvg;
  serializeRoundTripJson: typeof serializePptxRoundTripJson;
  writeRoundTrip: typeof writePptxRoundTrip;
}

const DEFAULT_OPERATIONS: OakitCliOperations = {
  parseRoundTripJson: parsePptxRoundTripJson,
  parsePptx: parsePptxWithDiagnostics,
  readRoundTrip: readPptxRoundTrip,
  replaceRoundTripText: replacePptxRoundTripText,
  renderPng: renderPptxToPng,
  renderSvg: renderPptxToSvg,
  serializeRoundTripJson: serializePptxRoundTripJson,
  writeRoundTrip: writePptxRoundTrip,
};

class CliUsageError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

const HELP = `Usage: oakit [convert] <input.pptx|-> [options]
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

function optionValue(iterator: Iterator<string>, option: string): string {
  const next = iterator.next();
  if (next.done) {
    throw new CliUsageError(
      'missing-option-value',
      `Option ${option} requires a value`,
    );
  }
  const value = next.value;
  if (value.startsWith('-') && value !== '-') {
    throw new CliUsageError(
      'missing-option-value',
      `Option ${option} requires a value`,
    );
  }
  return value;
}

function normalizeFormat(value: string): OakitCliFormat {
  if (value.toLowerCase() !== 'pptx') {
    throw new CliUsageError(
      'unsupported-format',
      `Unsupported Office format: ${value}`,
    );
  }
  return 'pptx';
}

function inferFormat(input: string, explicitFormat?: string): OakitCliFormat {
  if (explicitFormat !== undefined) return normalizeFormat(explicitFormat);
  if (input === '-') {
    throw new CliUsageError(
      'format-required',
      'Reading stdin requires --format pptx',
    );
  }
  return normalizeFormat(path.extname(input).slice(1));
}

function renderFormat(value: string): OakitRenderFormat {
  if (value !== 'png' && value !== 'svg') {
    throw new CliUsageError(
      'invalid-render-format',
      `Unsupported render format: ${value}`,
    );
  }
  return value;
}

function renderSlideNumbers(value: string): readonly number[] {
  const tokens = value.split(',');
  if (tokens.some((token) => !/^[1-9]\d*$/.test(token))) {
    throw new CliUsageError(
      'invalid-slides',
      'Render slides must be unique positive safe integers',
    );
  }
  const numbers = tokens.map((token) => Number(token));
  if (
    numbers.some((number) => !Number.isSafeInteger(number)) ||
    new Set(numbers).size !== numbers.length
  ) {
    throw new CliUsageError(
      'invalid-slides',
      'Render slides must be unique positive safe integers',
    );
  }
  return numbers;
}

function renderScale(value: string): number {
  if (!/^(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value)) {
    throw new CliUsageError(
      'invalid-scale',
      'Render scale must be a positive finite number',
    );
  }
  const scale = Number(value);
  if (!Number.isFinite(scale) || scale <= 0) {
    throw new CliUsageError(
      'invalid-scale',
      'Render scale must be a positive finite number',
    );
  }
  return scale;
}

function parseCommand(args: readonly string[]): CliCommand {
  if (args.includes('--help') || args.includes('-h')) return { action: 'help' };
  if (args.includes('--version') || args.includes('-v')) {
    return { action: 'version' };
  }

  const action =
    args[0] === 'edit-text' ||
    args[0] === 'render' ||
    args[0] === 'restore' ||
    args[0] === 'snapshot'
      ? args[0]
      : 'convert';
  const values =
    args[0] === 'convert' ||
    args[0] === 'edit-text' ||
    args[0] === 'render' ||
    args[0] === 'restore' ||
    args[0] === 'snapshot'
      ? args.slice(1)
      : [...args];
  const inputs: string[] = [];
  let documentOnly = false;
  let explicitFormat: string | undefined;
  let imageMode: 'base64' | 'none' = 'none';
  let output: string | undefined;
  let pretty = false;
  let selectedRenderFormat: OakitRenderFormat = 'png';
  let scale = 1;
  let slideNumbers: readonly number[] | undefined;
  let strict = false;
  let targetKey: string | undefined;
  let textValue: string | undefined;

  const iterator = values.values();
  for (const value of iterator) {
    if (value === '--document-only' && action === 'convert') {
      documentOnly = true;
    } else if (value === '--strict' && action === 'convert') {
      strict = true;
    } else if (
      value === '--pretty' &&
      (action === 'convert' || action === 'edit-text' || action === 'snapshot')
    ) {
      pretty = true;
    } else if (value === '--output' || value === '-o') {
      output = optionValue(iterator, value);
    } else if (
      value === '--format' &&
      action !== 'edit-text' &&
      action !== 'restore'
    ) {
      explicitFormat = optionValue(iterator, value);
    } else if (value === '--target' && action === 'edit-text') {
      targetKey = optionValue(iterator, value);
    } else if (value.startsWith('--target=') && action === 'edit-text') {
      targetKey = value.slice('--target='.length);
    } else if (value === '--value' && action === 'edit-text') {
      textValue = optionValue(iterator, value);
    } else if (value.startsWith('--value=') && action === 'edit-text') {
      textValue = value.slice('--value='.length);
    } else if (value === '--image-mode' && action === 'convert') {
      const selectedMode = optionValue(iterator, value);
      if (selectedMode !== 'none' && selectedMode !== 'base64') {
        throw new CliUsageError(
          'invalid-image-mode',
          `Unsupported image mode: ${selectedMode}`,
        );
      }
      imageMode = selectedMode;
    } else if (value === '--render-format' && action === 'render') {
      selectedRenderFormat = renderFormat(optionValue(iterator, value));
    } else if (value === '--slides' && action === 'render') {
      slideNumbers = renderSlideNumbers(optionValue(iterator, value));
    } else if (value === '--scale' && action === 'render') {
      scale = renderScale(optionValue(iterator, value));
    } else if (value === '-' || !value.startsWith('-')) {
      inputs.push(value);
    } else {
      throw new CliUsageError('unknown-option', `Unknown option: ${value}`);
    }
  }

  const [input, ...additionalInputs] = inputs;
  if (input === undefined) {
    throw new CliUsageError(
      'input-required',
      action === 'restore'
        ? 'A portable JSON input path is required'
        : action === 'edit-text'
          ? 'A portable JSON input path is required'
          : 'A PPTX input path is required',
    );
  }
  if (additionalInputs.length > 0) {
    throw new CliUsageError(
      'too-many-inputs',
      'Only one input document can be converted at a time',
    );
  }

  if (
    input !== '-' &&
    output !== undefined &&
    output !== '-' &&
    path.resolve(input) === path.resolve(output)
  ) {
    throw new CliUsageError(
      'output-overwrites-input',
      action === 'render'
        ? 'The render output directory must not overwrite the input document'
        : action === 'restore'
          ? 'The PowerPoint output path must not overwrite the portable JSON input'
          : action === 'edit-text'
            ? 'The JSON output path must not overwrite the portable JSON input'
            : 'The JSON output path must not overwrite the input document',
    );
  }

  if (action === 'restore') {
    if (output === undefined || output === '-') {
      throw new CliUsageError(
        'restore-output-required',
        'Restoring requires a PowerPoint output file',
      );
    }
    return { action, input, output };
  }

  if (action === 'edit-text') {
    if (targetKey === undefined || targetKey.length === 0) {
      throw new CliUsageError(
        'edit-target-required',
        'Editing text requires a non-empty --target run key',
      );
    }
    if (textValue === undefined) {
      throw new CliUsageError(
        'edit-value-required',
        'Editing text requires --value',
      );
    }
    return {
      action,
      input,
      ...(output === undefined ? {} : { output }),
      pretty,
      targetKey,
      value: textValue,
    };
  }

  const format = inferFormat(input, explicitFormat);
  if (action === 'render') {
    if (output === undefined || output === '-') {
      throw new CliUsageError(
        'render-output-required',
        'Rendering requires an output directory',
      );
    }
    return {
      action,
      format,
      input,
      output,
      renderFormat: selectedRenderFormat,
      scale,
      ...(slideNumbers === undefined ? {} : { slideNumbers }),
    };
  }

  if (action === 'snapshot') {
    return {
      action,
      format,
      input,
      ...(output === undefined ? {} : { output }),
      pretty,
    };
  }

  return {
    action: 'convert',
    documentOnly,
    format,
    imageMode,
    input,
    ...(output === undefined ? {} : { output }),
    pretty,
    strict,
  };
}

function errorJson(
  code: string,
  message: string,
  diagnostic?:
    | PptxDiagnostic
    | {
        actual: number;
        limit: number;
        limitName: string;
      },
): string {
  const error: {
    code: string;
    diagnostic?:
      PptxDiagnostic | { actual: number; limit: number; limitName: string };
    message: string;
  } = { code, message };
  if (diagnostic !== undefined) error.diagnostic = diagnostic;
  return `${JSON.stringify({ error })}\n`;
}

async function convert(
  command: ConvertCommand,
  io: OakitCliIo,
  operations: OakitCliOperations,
): Promise<number> {
  let input: Uint8Array;
  try {
    input =
      command.input === '-'
        ? await io.readStdin()
        : await io.readFile(command.input);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('input-read-failed', message));
    return 1;
  }

  let json: string;
  try {
    const result = await operations.parsePptx(input, {
      audioMode: 'none',
      errorMode: command.strict ? 'strict' : 'tolerant',
      imageMode: command.imageMode,
      videoMode: 'none',
    });
    const payload = command.documentOnly
      ? result.document
      : { format: command.format, ...result };
    json = `${JSON.stringify(payload, null, command.pretty ? 2 : 0)}\n`;
  } catch (error) {
    if (error instanceof PptxParseError) {
      io.writeStderr(
        errorJson(error.diagnostic.code, error.message, error.diagnostic),
      );
    } else {
      const message = error instanceof Error ? error.message : String(error);
      io.writeStderr(errorJson('conversion-failed', message));
    }
    return 1;
  }

  try {
    if (command.output === undefined || command.output === '-') {
      io.writeStdout(json);
    } else {
      await io.writeFile(command.output, json);
    }
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('output-write-failed', message));
    return 1;
  }
}

async function renderSlides(
  command: RenderCommand,
  io: OakitCliIo,
  operations: OakitCliOperations,
): Promise<number> {
  let input: Uint8Array;
  try {
    input =
      command.input === '-'
        ? await io.readStdin()
        : await io.readFile(command.input);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('input-read-failed', message));
    return 1;
  }

  let result:
    | Awaited<ReturnType<OakitCliOperations['renderPng']>>
    | Awaited<ReturnType<OakitCliOperations['renderSvg']>>;
  try {
    const options = {
      scale: command.scale,
      ...(command.slideNumbers === undefined
        ? {}
        : { slideNumbers: command.slideNumbers }),
    };
    result =
      command.renderFormat === 'png'
        ? await operations.renderPng(input, options)
        : await operations.renderSvg(input, options);
  } catch (error) {
    if (error instanceof PptxParseError) {
      io.writeStderr(
        errorJson(error.diagnostic.code, error.message, error.diagnostic),
      );
    } else if (error instanceof PptxRenderError) {
      io.writeStderr(errorJson(error.code, error.message));
    } else {
      const message = error instanceof Error ? error.message : String(error);
      io.writeStderr(errorJson('render-failed', message));
    }
    return 1;
  }

  const slides = result.slides.map((slide) => {
    const filename = `slide-${slide.slideNumber}.${slide.format}`;
    return {
      data: slide.data,
      manifest: {
        byteLength: slide.data.byteLength,
        file: filename,
        format: slide.format,
        height: slide.height,
        mimeType: slide.mimeType,
        slideNumber: slide.slideNumber,
        warnings: slide.warnings,
        width: slide.width,
      },
    };
  });

  try {
    await io.createDirectory(command.output);
    for (const slide of slides) {
      await io.writeBinaryFile(
        path.join(command.output, slide.manifest.file),
        slide.data,
      );
    }
    await io.writeFile(
      path.join(command.output, 'manifest.json'),
      `${JSON.stringify(
        {
          format: 'pptx-render',
          renderFormat: command.renderFormat,
          scale: command.scale,
          slides: slides.map(({ manifest }) => manifest),
          source: command.input === '-' ? 'stdin' : command.input,
        },
        null,
        2,
      )}\n`,
    );
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('output-write-failed', message));
    return 1;
  }
}

function portableLimitJson(error: PptxRoundTripPortableLimitError): string {
  return errorJson('portable-limit-exceeded', error.message, {
    actual: error.actual,
    limit: error.limit,
    limitName: error.limitName,
  });
}

async function snapshotPortableJson(
  command: SnapshotCommand,
  io: OakitCliIo,
  operations: OakitCliOperations,
): Promise<number> {
  let input: Uint8Array;
  try {
    input =
      command.input === '-'
        ? await io.readStdin()
        : await io.readFile(command.input);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('input-read-failed', message));
    return 1;
  }

  let json: string;
  try {
    const runtime = await operations.readRoundTrip(input);
    const portable = await operations.serializeRoundTripJson(runtime);
    json = `${JSON.stringify(portable, null, command.pretty ? 2 : undefined)}\n`;
  } catch (error) {
    if (error instanceof PptxParseError) {
      io.writeStderr(
        errorJson(error.diagnostic.code, error.message, error.diagnostic),
      );
    } else if (error instanceof PptxRoundTripPortableLimitError) {
      io.writeStderr(portableLimitJson(error));
    } else if (error instanceof PptxWriteError) {
      io.writeStderr(errorJson(error.code, error.message));
    } else {
      const message = error instanceof Error ? error.message : String(error);
      io.writeStderr(errorJson('snapshot-failed', message));
    }
    return 1;
  }

  try {
    if (command.output === undefined || command.output === '-') {
      io.writeStdout(json);
    } else {
      await io.writeFile(command.output, json);
    }
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('output-write-failed', message));
    return 1;
  }
}

type PortableJsonInput = { ok: false } | { ok: true; value: unknown };

async function readPortableJsonInput(
  inputPath: string,
  io: OakitCliIo,
): Promise<PortableJsonInput> {
  let input: Uint8Array;
  try {
    input =
      inputPath === '-' ? await io.readStdin() : await io.readFile(inputPath);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('input-read-failed', message));
    return { ok: false };
  }

  let source: string;
  try {
    source = new TextDecoder('utf-8', { fatal: true }).decode(input);
  } catch {
    io.writeStderr(
      errorJson(
        'invalid-portable-json',
        'Portable JSON input must be valid UTF-8',
      ),
    );
    return { ok: false };
  }

  try {
    return { ok: true, value: JSON.parse(source) as unknown };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('invalid-portable-json', message));
    return { ok: false };
  }
}

async function editPortableText(
  command: EditTextCommand,
  io: OakitCliIo,
  operations: OakitCliOperations,
): Promise<number> {
  const input = await readPortableJsonInput(command.input, io);
  if (!input.ok) return 1;

  let json: string;
  try {
    const runtime = await operations.parseRoundTripJson(input.value);
    const edited = await operations.replaceRoundTripText(runtime, {
      targetKey: command.targetKey,
      value: command.value,
    });
    const portable = await operations.serializeRoundTripJson(edited);
    json = `${JSON.stringify(portable, null, command.pretty ? 2 : undefined)}\n`;
  } catch (error) {
    if (error instanceof PptxRoundTripPortableLimitError) {
      io.writeStderr(portableLimitJson(error));
    } else if (error instanceof PptxWriteError) {
      io.writeStderr(errorJson(error.code, error.message));
    } else {
      const message = error instanceof Error ? error.message : String(error);
      io.writeStderr(errorJson('edit-text-failed', message));
    }
    return 1;
  }

  try {
    if (command.output === undefined || command.output === '-') {
      io.writeStdout(json);
    } else {
      await io.writeFile(command.output, json);
    }
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('output-write-failed', message));
    return 1;
  }
}

async function restorePortableJson(
  command: RestoreCommand,
  io: OakitCliIo,
  operations: OakitCliOperations,
): Promise<number> {
  const input = await readPortableJsonInput(command.input, io);
  if (!input.ok) return 1;

  let output: Uint8Array;
  try {
    const runtime = await operations.parseRoundTripJson(input.value);
    const result = await operations.writeRoundTrip(runtime);
    output = result.data;
  } catch (error) {
    if (error instanceof PptxRoundTripPortableLimitError) {
      io.writeStderr(portableLimitJson(error));
    } else if (error instanceof PptxWriteError) {
      io.writeStderr(errorJson(error.code, error.message));
    } else {
      const message = error instanceof Error ? error.message : String(error);
      io.writeStderr(errorJson('restore-failed', message));
    }
    return 1;
  }

  try {
    await io.writeBinaryFile(command.output, output);
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    io.writeStderr(errorJson('output-write-failed', message));
    return 1;
  }
}

export async function runOakitCli(
  args: readonly string[],
  io: OakitCliIo,
  version: string,
  operationOverrides: Partial<OakitCliOperations> = {},
): Promise<number> {
  const operations = { ...DEFAULT_OPERATIONS, ...operationOverrides };
  let command: CliCommand;
  try {
    command = parseCommand(args);
  } catch (error) {
    if (!(error instanceof CliUsageError)) throw error;
    io.writeStderr(errorJson(error.code, error.message));
    return 2;
  }

  if (command.action === 'help') {
    io.writeStdout(HELP);
    return 0;
  }
  if (command.action === 'version') {
    io.writeStdout(`oakit ${version}\n`);
    return 0;
  }
  if (command.action === 'render') {
    return renderSlides(command, io, operations);
  }
  if (command.action === 'edit-text') {
    return editPortableText(command, io, operations);
  }
  if (command.action === 'restore') {
    return restorePortableJson(command, io, operations);
  }
  if (command.action === 'snapshot') {
    return snapshotPortableJson(command, io, operations);
  }
  return convert(command, io, operations);
}
