# office2json

The communication layer between AI and Microsoft Office documents.

`office2json` is planned as a bidirectional TypeScript library for lossless
conversion between JSON and:

- PowerPoint (`.pptx`)
- Excel (`.xlsx`)
- Word (`.docx`)

## Status

The first parser milestone supports PowerPoint (`.pptx`) to typed JSON,
including slide backgrounds, layouts, masters, text, shapes, images, media,
tables, charts, diagrams, math, notes, and transitions. Fidelity work will
continue fixture by fixture; reverse conversion and the Excel/Word formats are
not implemented yet.

## Usage

```ts
import { parsePptx } from 'office2json';

const document = await parsePptx(arrayBuffer, {
  imageMode: 'base64',
  videoMode: 'none',
  audioMode: 'none',
});
```

The same API is available from the format-specific entry point:

```ts
import { parsePptx } from 'office2json/pptx';
```

See [`docs/architecture.md`](docs/architecture.md) for module boundaries and
the migration strategy.

## Development

```bash
pnpm install
pnpm check
```

## License

MIT © EvoElsewhere.
