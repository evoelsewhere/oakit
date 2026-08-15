<p align="center">
  <img src="docs/assets/oakit-logo.png" alt="OAKit logo" width="220" />
</p>

<h1 align="center">Office Agent Kit</h1>

<p align="center">
  <strong> Office documents become agent-ready knowledge. </strong>
</p>

OAKit gives agents and automation systems a reliable way to read, understand,
edit, and eventually generate PowerPoint, Excel, and Word documents through a
consistent structured model. It owns the difficult OOXML work—ZIP packages,
relationships, inheritance, units, media, malformed input, and producer
differences—so agent workflows can operate on meaningful document data instead
of raw XML.

> **Project status:** pre-stable (`0.0.0`). The implemented public capability is
> PowerPoint (`.pptx`) reading. Excel, Word, and document writing are product
> direction, not completed APIs yet.

## Why OAKit

Office files are not single documents internally. They are ZIP packages made
of interconnected XML parts, relationships, themes, layouts, media, charts,
and vendor-specific extensions. That representation is a poor tool boundary
for an AI agent.

OAKit turns those internals into bounded, deterministic application data that is
easier to:

- summarize, classify, index, and search;
- inspect slide structure and extract semantic content;
- build document-aware tools and agent actions;
- validate generated changes before writing a file;
- run consistently in Node.js and modern browsers;
- process untrusted uploads with explicit diagnostics and resource limits.

OAKit is model- and framework-neutral. It does not require a particular LLM,
agent runtime, tool-calling protocol, or vector database.

## Format support

| Format               | Read | Write | Status                      |
| -------------------- | ---- | ----- | --------------------------- |
| PowerPoint (`.pptx`) | Yes  | No    | Active fidelity development |
| Excel (`.xlsx`)      | No   | No    | Planned                     |
| Word (`.docx`)       | No   | No    | Planned                     |

The PowerPoint reader currently handles:

- slide order, size, backgrounds, layouts, masters, themes, and notes;
- rich text, paragraphs, lists, links, fonts, and text fitting;
- preset and custom shapes, connectors, fills, borders, shadows, and groups;
- images, cropping, filters, embedded audio, and embedded video;
- tables, charts, SmartArt diagrams, and Office Math;
- transitions, relationship resolution, diagnostics, and resource limits.

OOXML has a very large extension surface. Unsupported optional structures may
be omitted with a diagnostic rather than represented inaccurately.

## Installation

The target npm package is `@evoelsewhere/oakit`:

```bash
pnpm add @evoelsewhere/oakit
```

```bash
npm install @evoelsewhere/oakit
```

### Homebrew CLI

On macOS or Linux, install the command-line interface from the EvoElsewhere
tap after the first formula release:

```bash
brew install evoelsewhere/tap/oakit
```

The Homebrew formula installs the `oakit` executable. Use npm or pnpm when the
programmatic JavaScript API is required.

The npm package and Homebrew formula have not been published yet. Until then,
use the repository directly for development.

## Command-line interface

The package installs the `oakit` executable for deterministic Office-to-JSON
workflows in terminals, scripts, CI jobs, and agent sandboxes. The current CLI
accepts PowerPoint (`.pptx`) input.

Install it globally after the first npm release:

```bash
npm install --global @evoelsewhere/oakit
oakit --version
```

It can also be run without a global installation:

```bash
npx --package @evoelsewhere/oakit oakit deck.pptx --pretty
```

### Convert a file

JSON is written to stdout by default, making the command suitable for pipes:

```bash
oakit deck.pptx > deck.json
```

Use the explicit `convert` command and `--output` when writing a file directly:

```bash
oakit convert deck.pptx --output deck.json --pretty
```

Both command forms are equivalent. OAKit refuses to overwrite the input document
with JSON output.

### Read from stdin

Use `-` as the input path and provide the format explicitly:

```bash
cat deck.pptx | oakit - --format pptx --document-only > deck.json
```

`--format pptx` is required for stdin because there is no filename extension
from which to infer the format.

### CLI options

```text
Usage: oakit [convert] <input.pptx|-> [options]

Options:
  -o, --output <file>          Write JSON to a file instead of stdout
      --format <pptx>          Input format; required when reading stdin
      --strict                 Reject malformed optional OOXML content
      --pretty                 Format JSON with two-space indentation
      --document-only          Omit format metadata and diagnostics
      --image-mode <mode>      Image output: none (default) or base64
  -h, --help                   Show help
  -v, --version                Show the installed OAKit version
```

The default output is an envelope that keeps format and recovery information
available to automation:

```json
{
  "format": "pptx",
  "document": {
    "slides": []
  },
  "diagnostics": []
}
```

Use `--document-only` when a downstream tool accepts only the normalized
document model. Use `--strict` when partial recovery is not acceptable. Images
are omitted by default to keep agent context and pipeline output bounded;
enable `--image-mode base64` only when the binary representation is required.
Audio and video payloads are never emitted by the CLI.

### Errors and exit codes

Errors are written as single-line JSON to stderr without a stack trace:

```json
{
  "error": {
    "code": "unsupported-format",
    "message": "Unsupported Office format: docx"
  }
}
```

| Exit code | Meaning                                         |
| --------- | ----------------------------------------------- |
| `0`       | Conversion, help, or version completed normally |
| `1`       | Input read or document conversion failed        |
| `2`       | Invalid command-line usage                      |

The CLI processes one document per invocation. Resource-limit failures remain
fatal in tolerant mode, matching the programmatic API's security boundary.

## Quick start

### Node.js

```ts
import { readFile } from 'node:fs/promises';
import { parsePptxWithDiagnostics } from '@evoelsewhere/oakit';

const input = await readFile('./quarterly-review.pptx');
const { document, diagnostics } = await parsePptxWithDiagnostics(input, {
  imageMode: 'none',
  errorMode: 'tolerant',
});

console.log({
  slideCount: document.slides.length,
  size: document.size,
  fonts: document.usedFonts,
  diagnostics,
});
```

Node.js `Buffer` extends `Uint8Array`, so bytes returned by `readFile` can be
passed directly to OAKit.

### Browser

```ts
import { parsePptx } from '@evoelsewhere/oakit/pptx';

const picker = document.querySelector<HTMLInputElement>('#presentation');
const file = picker?.files?.[0];

if (file) {
  const presentation = await parsePptx(file, {
    imageMode: 'both',
    videoMode: 'blob',
    audioMode: 'blob',
  });

  console.log(presentation.slides);
}
```

### Agent tool boundary

Use the diagnostic API when an agent must distinguish usable partial output
from a clean parse:

```ts
import { parsePptxWithDiagnostics } from '@evoelsewhere/oakit';

export async function inspectPresentation(bytes: Uint8Array) {
  const result = await parsePptxWithDiagnostics(bytes, {
    imageMode: 'none',
    videoMode: 'none',
    audioMode: 'none',
    errorMode: 'tolerant',
  });

  return {
    kind: 'presentation' as const,
    document: result.document,
    diagnostics: result.diagnostics,
  };
}
```

Document text is untrusted content. An agent host must keep it in the data
portion of its prompt or tool result and must not treat instructions embedded
in a document as trusted system or developer instructions.

## Public PowerPoint API

Both entry points expose the same reader:

```ts
import {
  parsePptx,
  parsePptxWithDiagnostics,
  PptxParseError,
} from '@evoelsewhere/oakit';

import { parsePptx as parsePptxFormat } from '@evoelsewhere/oakit/pptx';
```

The format-specific entry point is preferred when an application only needs
PowerPoint support.

### Input

```ts
type PptxInput = ArrayBuffer | Uint8Array | Blob;
```

### Options

```ts
interface PptxParseOptions {
  imageMode?: 'base64' | 'blob' | 'both' | 'none';
  videoMode?: 'blob' | 'none';
  audioMode?: 'blob' | 'none';
  errorMode?: 'tolerant' | 'strict';
  limits?: PptxResourceLimits;
}
```

| Option      | Default       | Behavior                                            |
| ----------- | ------------- | --------------------------------------------------- |
| `imageMode` | `base64`      | Return data URLs, object URLs, both, or neither.    |
| `videoMode` | `none`        | Create object URLs for supported embedded video.    |
| `audioMode` | `none`        | Create object URLs for supported embedded audio.    |
| `errorMode` | `tolerant`    | Recover with diagnostics or reject malformed OOXML. |
| `limits`    | Safe defaults | Bound archive, XML, media, and slide processing.    |

### Output

```ts
interface PptxDocument {
  size: {
    width: number;
    height: number;
  };
  themeColors: string[];
  usedFonts: string[];
  slides: PptxSlide[];
}
```

Positions and dimensions use points. Each slide separates authored elements
from inherited layout and master content:

```ts
interface PptxSlide {
  fill: Fill;
  elements: PptxElement[];
  layoutElements: PptxElement[];
  note: string;
  transition?: SlideTransition | null;
}
```

Elements use a discriminated `type` field:

| `type`    | Content                                                       |
| --------- | ------------------------------------------------------------- |
| `text`    | Positioned rich-text HTML and text layout                     |
| `shape`   | Shape metadata and an SVG-compatible path when available      |
| `image`   | Package reference, selected representation, crop, and filters |
| `video`   | Embedded or external reference and optional object URL        |
| `audio`   | Embedded reference and optional object URL                    |
| `table`   | Cells, merges, dimensions, fills, and borders                 |
| `chart`   | Normalized series, labels, colors, and options                |
| `diagram` | SmartArt drawing elements and logical text                    |
| `math`    | Parsed LaTeX and an optional fallback image                   |
| `group`   | Nested elements in the group coordinate space                 |

Text is returned as an escaped HTML fragment. Applications that inject
document HTML into a page should still apply their own sanitizer as defense in
depth.

## Diagnostics and strict mode

`parsePptx` returns the document directly. `parsePptxWithDiagnostics` returns:

```ts
interface PptxParseResult {
  document: PptxDocument;
  diagnostics: PptxDiagnostic[];
}
```

Tolerant mode may omit malformed optional content while recording a structured
diagnostic. Strict mode rejects malformed XML, unsafe relationships, invalid
values, and missing required parts with `PptxParseError`.

Resource-limit violations always reject, including in tolerant mode. They
represent a security boundary, not a recoverable fidelity problem.

## Security model

OAKit treats every uploaded package as untrusted input. The reader:

- rejects unsafe package and relationship paths;
- rejects malformed XML structures and forbidden declarations;
- validates numeric values before conversion;
- does not execute macros, scripts, media, or hyperlinks;
- does not fetch external relationships;
- filters supported hyperlink protocols;
- bounds compressed input, ZIP entries, expanded bytes, XML complexity,
  embedded media, and slide count.

Default limits:

| Limit                     |   Default |
| ------------------------- | --------: |
| Compressed input          |   100 MiB |
| Non-directory ZIP entries |    10,000 |
| Total declared expansion  |   256 MiB |
| One expanded package part |    64 MiB |
| One expanded XML part     |    16 MiB |
| XML nesting depth         |       128 |
| XML elements per part     |   250,000 |
| XML elements per package  | 1,000,000 |
| One expanded media part   |    64 MiB |
| Slides                    |     1,000 |

For public uploads, also isolate parsing in a worker or process and enforce an
outer timeout and memory limit.

## Media lifecycle

When a blob mode is enabled, OAKit calls `URL.createObjectURL`. The caller owns
the returned URLs and must release them:

```ts
URL.revokeObjectURL(element.blob);
```

Remember to traverse nested group and diagram elements when releasing media.

## Runtime support

- Node.js 20, 22, and 24;
- modern browsers with `Blob` and `URL.createObjectURL` support;
- ESM and CommonJS;
- declarations and source maps.

## Development

```bash
pnpm install
pnpm check
```

| Command                  | Purpose                                                 |
| ------------------------ | ------------------------------------------------------- |
| `pnpm dev`               | Rebuild in watch mode                                   |
| `pnpm test`              | Run deterministic unit, integration, and property tests |
| `pnpm test:browser`      | Run the public API suite in Chromium                    |
| `pnpm test:corpus`       | Verify PowerPoint and LibreOffice documents             |
| `pnpm test:corpus:large` | Include the large Google Slides corpus                  |
| `pnpm test:mutation`     | Measure whether tests detect behavioral mutations       |
| `pnpm test:package`      | Smoke-test package exports and the bundled CLI          |
| `pnpm typecheck`         | Run strict type checking                                |
| `pnpm lint`              | Run ESLint                                              |
| `pnpm format:check`      | Verify formatting                                       |
| `pnpm build`             | Build ESM, CommonJS, declarations, and source maps      |
| `pnpm check`             | Run the required pull-request quality gates             |

The fast CI matrix runs on Node.js 20, 22, and 24 plus Chromium. Producer
corpus and mutation suites run in the reliability workflow.

Read [docs/architecture.md](docs/architecture.md) before changing parser
ownership, public models, resource handling, or format boundaries. Development
rules for coding agents and contributors live in [AGENTS.md](AGENTS.md).

## Roadmap

- stabilize the normalized PowerPoint model;
- expand real-producer fidelity and adversarial corpus coverage;
- add a mutation-tested PowerPoint writer;
- introduce Excel and Word as isolated format domains;
- expose higher-level document operations suitable for agent tools;
- keep the core independent of model vendors and agent frameworks.

Capabilities are documented only after their public API, implementation, and
tests exist.

## License

[MIT](LICENSE) © 2026 EvoElsewhere.
