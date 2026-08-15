import path from 'node:path';

import {
  parsePptxWithDiagnostics,
  PptxParseError,
  type PptxDiagnostic,
} from '../formats/pptx';

export interface OakitCliIo {
  readFile(filename: string): Promise<Uint8Array>;
  readStdin(): Promise<Uint8Array>;
  writeFile(filename: string, value: string): Promise<void>;
  writeStderr(value: string): void;
  writeStdout(value: string): void;
}

type OakitCliFormat = 'pptx';

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

type CliCommand = ConvertCommand | { action: 'help' } | { action: 'version' };

class CliUsageError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

const HELP = `Usage: oakit [convert] <input.pptx|-> [options]

Convert a PowerPoint Open XML presentation into deterministic JSON.

Options:
  -o, --output <file>          Write JSON to a file instead of stdout
      --format <pptx>          Input format; required when reading stdin
      --strict                 Reject malformed optional OOXML content
      --pretty                 Format JSON with two-space indentation
      --document-only          Omit format metadata and diagnostics
      --image-mode <mode>      Image output: none (default) or base64
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

function parseCommand(args: readonly string[]): CliCommand {
  if (args.includes('--help') || args.includes('-h')) return { action: 'help' };
  if (args.includes('--version') || args.includes('-v')) {
    return { action: 'version' };
  }

  const values = args[0] === 'convert' ? args.slice(1) : [...args];
  const inputs: string[] = [];
  let documentOnly = false;
  let explicitFormat: string | undefined;
  let imageMode: 'base64' | 'none' = 'none';
  let output: string | undefined;
  let pretty = false;
  let strict = false;

  const iterator = values.values();
  for (const value of iterator) {
    if (value === '--document-only') {
      documentOnly = true;
    } else if (value === '--strict') {
      strict = true;
    } else if (value === '--pretty') {
      pretty = true;
    } else if (value === '--output' || value === '-o') {
      output = optionValue(iterator, value);
    } else if (value === '--format') {
      explicitFormat = optionValue(iterator, value);
    } else if (value === '--image-mode') {
      const selectedMode = optionValue(iterator, value);
      if (selectedMode !== 'none' && selectedMode !== 'base64') {
        throw new CliUsageError(
          'invalid-image-mode',
          `Unsupported image mode: ${selectedMode}`,
        );
      }
      imageMode = selectedMode;
    } else if (value === '-' || !value.startsWith('-')) {
      inputs.push(value);
    } else {
      throw new CliUsageError('unknown-option', `Unknown option: ${value}`);
    }
  }

  const [input, ...additionalInputs] = inputs;
  if (input === undefined) {
    throw new CliUsageError('input-required', 'A PPTX input path is required');
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
      'The JSON output path must not overwrite the input document',
    );
  }

  return {
    action: 'convert',
    documentOnly,
    format: inferFormat(input, explicitFormat),
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
  diagnostic?: PptxDiagnostic,
): string {
  const error: { code: string; diagnostic?: PptxDiagnostic; message: string } =
    {
      code,
      message,
    };
  if (diagnostic !== undefined) error.diagnostic = diagnostic;
  return `${JSON.stringify({ error })}\n`;
}

async function convert(
  command: ConvertCommand,
  io: OakitCliIo,
  parsePptx: typeof parsePptxWithDiagnostics,
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
    const result = await parsePptx(input, {
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

export async function runOakitCli(
  args: readonly string[],
  io: OakitCliIo,
  version: string,
  parsePptx: typeof parsePptxWithDiagnostics = parsePptxWithDiagnostics,
): Promise<number> {
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
  return convert(command, io, parsePptx);
}
