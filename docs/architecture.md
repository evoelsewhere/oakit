# Architecture

The source tree separates reusable Office Open XML infrastructure from each
document format:

```text
src/
├── common/                 Shared binary, media, text, number, and XML helpers
└── formats/
    └── pptx/
        ├── index.ts        Public PowerPoint entry point
        ├── parser.ts       Package and slide orchestration
        ├── types.ts        Public PowerPoint JSON model
        └── internal/       Domain parsers for shapes, text, fills, charts, etc.
```

Future `xlsx` and `docx` implementations should be sibling directories under
`src/formats`. Code belongs in `src/common` only when it is independent of a
specific Office format and can be tested through a stable shared contract.

## PowerPoint parsing flow

1. `parsePptx` opens the OPC ZIP package and reads presentation metadata.
2. The parser resolves slide, layout, master, theme, media, and diagram
   relationships.
3. Each DrawingML domain parser converts its OOXML subtree into the public JSON
   model.
4. The root parser returns document size, fonts, theme colors, slides,
   transitions, notes, and slide elements.

The compatibility port is compiled under strict TypeScript, including
`noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`. New behavior
should be introduced with schema-focused types and fixture tests instead of
loosening compiler or lint rules.

## Direction support

The current implementation supports `.pptx` to JSON. JSON to `.pptx`, both
directions for `.xlsx`, and both directions for `.docx` remain separate future
format milestones.
