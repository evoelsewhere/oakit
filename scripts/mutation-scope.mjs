export const mutatedFiles = [
  'src/cli/node-io.ts',
  'src/cli/run.ts',
  'src/common/archive/read-entry.ts',
  'src/common/binary/base64.ts',
  'src/common/media/media-type.ts',
  'src/common/numbers.ts',
  'src/common/ooxml/units.ts',
  'src/common/opc/part-uri.ts',
  'src/common/text/css.ts',
  'src/common/text/html.ts',
  'src/common/xml/normalize.ts',
  'src/common/xml/read-xml.ts',
  'src/common/xml/tree.ts',
  'src/common/xml/types.ts',
  'src/common/xml/validate.ts',
  'src/formats/pptx/internal/resource-limits.ts',
];

export const pendingMutationFiles = [
  'src/cli.ts',
  'src/formats/pptx/errors.ts',
  'src/formats/pptx/index.ts',
  'src/formats/pptx/internal/animation.ts',
  'src/formats/pptx/internal/border.ts',
  'src/formats/pptx/internal/chart.ts',
  'src/formats/pptx/internal/color.ts',
  'src/formats/pptx/internal/context.ts',
  'src/formats/pptx/internal/diagram.ts',
  'src/formats/pptx/internal/fill.ts',
  'src/formats/pptx/internal/font-style.ts',
  'src/formats/pptx/internal/math.ts',
  'src/formats/pptx/internal/paragraph.ts',
  'src/formats/pptx/internal/position.ts',
  'src/formats/pptx/internal/scheme-color.ts',
  'src/formats/pptx/internal/shadow.ts',
  'src/formats/pptx/internal/shape-path.ts',
  'src/formats/pptx/internal/shape.ts',
  'src/formats/pptx/internal/table.ts',
  'src/formats/pptx/internal/text-insets.ts',
  'src/formats/pptx/internal/text.ts',
  'src/formats/pptx/internal/xml-reader.ts',
  'src/formats/pptx/parser.ts',
];

export const excludedMutationFiles = [
  {
    file: 'src/common/index.ts',
    reason: 'Pure re-export barrel with no runtime decisions or values.',
  },
  {
    file: 'src/formats/pptx/types.ts',
    reason: 'Compile-time-only interfaces and type aliases.',
  },
  {
    file: 'src/index.ts',
    reason: 'Pure public re-export barrel with no runtime decisions or values.',
  },
  {
    file: 'src/types/txml.d.ts',
    reason: 'Ambient dependency declaration with no emitted runtime code.',
  },
];
