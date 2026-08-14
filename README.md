# office2json

The typed communication layer between AI applications and Microsoft Office
documents.

`office2json` reads Office Open XML packages and turns document content into a
renderer-friendly TypeScript model. The current milestone parses PowerPoint
presentations (`.pptx`); Excel, Word, and reverse conversion are planned as
separate format milestones.

> **Project status:** early development (`0.0.0`). The PowerPoint parser is
> usable, but the data model may still change before the first stable release.

## Format support

| Format     | Office to JSON | JSON to Office | Status                      |
| ---------- | -------------- | -------------- | --------------------------- |
| PowerPoint | Yes            | No             | Active fidelity development |
| Excel      | No             | No             | Planned                     |
| Word       | No             | No             | Planned                     |

The PowerPoint parser currently handles:

- slide size, backgrounds, layouts, masters, themes, and speaker notes;
- text boxes, rich-text runs, paragraphs, lists, links, and text fitting;
- preset and custom shapes, connectors, fills, borders, shadows, and groups;
- images, image cropping and filters, embedded video, and embedded audio;
- tables, charts, SmartArt diagrams, and Office Math;
- slide transitions and document-order metadata.

Support is fixture-driven. Real-world OOXML can contain vendor extensions and
fallback branches that are not covered yet, so consumers should tolerate
missing optional fields and unsupported elements.

## Requirements

- Node.js 20 or newer, or a modern browser with `Blob` and
  `URL.createObjectURL` support;
- ESM or CommonJS;
- a complete `.pptx` package supplied as `ArrayBuffer`, `Uint8Array`, or
  `Blob`.

## Installation

```bash
pnpm add office2json
```

Equivalent npm and Yarn commands work as well.

## Quick start

### Node.js

Node.js `Buffer` extends `Uint8Array`, so a file read from disk can be passed
directly to the parser:

```ts
import { readFile } from 'node:fs/promises';
import { parsePptx } from 'office2json';

const input = await readFile('./deck.pptx');
const presentation = await parsePptx(input);

console.log(presentation.size);
console.log(presentation.slides[0]?.elements);
```

### Browser

```ts
import { parsePptx } from 'office2json/pptx';

const input = document.querySelector<HTMLInputElement>('#presentation');
const file = input?.files?.[0];

if (file) {
  const presentation = await parsePptx(file, {
    imageMode: 'both',
    videoMode: 'blob',
    audioMode: 'blob',
  });

  console.log(presentation.slides);
}
```

Both package entry points expose the same PowerPoint API:

```ts
import { parsePptx } from 'office2json';
import { parsePptx as parsePptxFormat } from 'office2json/pptx';
```

The format-specific entry point is preferred when an application only needs
PowerPoint support.

## Parser options

```ts
interface PptxParseOptions {
  imageMode?: 'base64' | 'blob' | 'both' | 'none';
  videoMode?: 'blob' | 'none';
  audioMode?: 'blob' | 'none';
}
```

| Option      | Default  | Behavior                                                     |
| ----------- | -------- | ------------------------------------------------------------ |
| `imageMode` | `base64` | Return images as data URLs, object URLs, both, or neither.   |
| `videoMode` | `none`   | Create object URLs for supported embedded video when `blob`. |
| `audioMode` | `none`   | Create object URLs for supported embedded audio when `blob`. |

Every media element keeps a `ref` to its package part or external target.
Disabled representations are returned as empty strings. Selecting `blob`
causes the parser to call `URL.createObjectURL`; the application owns those
URLs and must release them when they are no longer needed:

```ts
for (const slide of presentation.slides) {
  for (const element of slide.elements) {
    if ('blob' in element && element.blob) {
      URL.revokeObjectURL(element.blob);
    }
  }
}
```

Nested group and diagram elements should also be traversed when an application
enables blob output for content inside those containers.

## Returned document

`parsePptx` returns a `PptxDocument`:

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

- `size`, element positions, and element dimensions use **points**. There are
  72 points per inch.
- `themeColors` contains the available theme accent colors (`accent1` through
  `accent6`) as hexadecimal strings.
- `usedFonts` currently contains font faces declared in PowerPoint's embedded
  font list. It is not yet a complete scan of every text run.
- slides are returned in their numeric package order.

Each slide separates authored content from inherited decorative content:

```ts
interface PptxSlide {
  fill: Fill;
  elements: PptxElement[];
  layoutElements: PptxElement[];
  note: string;
  transition?: SlideTransition | null;
}
```

- `elements` contains content authored on the slide;
- `layoutElements` contains non-placeholder shapes inherited from the layout
  and master;
- `note` is an HTML fragment extracted from the body placeholder of the notes
  slide;
- `transition` is `null` when the slide, layout, and master define no supported
  transition.

### Element model

Elements use a discriminated `type` field:

| `type`    | Main payload                                                    |
| --------- | --------------------------------------------------------------- |
| `text`    | Positioned rich-text HTML, fill, border, and text layout.       |
| `shape`   | Shape metadata plus an SVG-compatible path when available.      |
| `image`   | Package reference, selected binary representation, crop/filter. |
| `video`   | Embedded or external media reference and optional object URL.   |
| `audio`   | Embedded media reference and optional object URL.               |
| `table`   | Cell matrix, merges, row heights, column widths, and borders.   |
| `chart`   | Normalized chart series, labels, colors, and chart options.     |
| `diagram` | SmartArt drawing elements and its logical text list.            |
| `math`    | Parsed LaTeX plus the fallback image when present.              |
| `group`   | Nested elements transformed into the group's coordinate space.  |

Text content is an HTML fragment rather than plain text. Treat it as untrusted
document input: sanitize it before injecting it into an HTML page when files
come from users you do not trust.

All public PowerPoint types are exported from both entry points:

```ts
import type {
  PptxDocument,
  PptxElement,
  PptxInput,
  PptxParseOptions,
  PptxSlide,
} from 'office2json/pptx';
```

## Error behavior

Invalid ZIP input causes `parsePptx` to reject. Missing optional OOXML parts
are generally treated as absent so that partially supported presentations can
still be parsed. Unsupported nodes may be skipped instead of producing a
placeholder element.

The parser does not currently enforce archive-size or decompression limits.
Applications accepting untrusted uploads should apply file-size, timeout, and
resource limits before parsing.

## Development

```bash
pnpm install
pnpm check
```

Useful commands:

| Command             | Purpose                                        |
| ------------------- | ---------------------------------------------- |
| `pnpm dev`          | Rebuild the package in watch mode.             |
| `pnpm test`         | Run the Vitest suite once.                     |
| `pnpm test:watch`   | Run tests in watch mode.                       |
| `pnpm typecheck`    | Check strict TypeScript types without output.  |
| `pnpm lint`         | Run type-aware ESLint rules.                   |
| `pnpm format:check` | Verify Prettier formatting.                    |
| `pnpm build`        | Build ESM, CommonJS, source maps, and typings. |
| `pnpm check`        | Run every required quality gate.               |

The test suite builds a minimal OOXML package in memory. New fidelity work
should add the smallest fixture that demonstrates the target OOXML structure
and assert the normalized public result.

## Architecture

Read [docs/architecture.md](docs/architecture.md) for the parsing pipeline,
OOXML relationship resolution, module boundaries, inheritance rules, media
lifecycle, extension workflow, and design constraints.

## Current limitations and roadmap

- The public model is pre-stable and may receive breaking corrections.
- PowerPoint output is normalized for consumption; it cannot currently be
  converted back into an equivalent `.pptx` package.
- Unsupported OOXML extensions and uncommon chart, SmartArt, animation, and
  shape variants require additional fixtures.
- Excel (`.xlsx`), Word (`.docx`), and every JSON-to-Office writer are not yet
  implemented.
- Large packages are loaded into memory by JSZip; streaming parsing is not yet
  available.

The long-term direction is one shared OOXML foundation with isolated format
adapters and explicit reader/writer capabilities. Fidelity is added one tested
document structure at a time.

## License

[MIT](LICENSE) © 2026 EvoElsewhere.
