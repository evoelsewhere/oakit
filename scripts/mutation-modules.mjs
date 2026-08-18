export const mutationModules = [
  {
    estimatedSeconds: 180,
    name: 'orchestration',
    source: 'src/formats/pptx/roundtrip/orchestration.ts',
    tests: [
      'test/pptx/roundtrip-orchestration.test.ts',
      'test/pptx/roundtrip-patch-text.test.ts',
      'test/pptx/roundtrip-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 150,
    name: 'package-preservation',
    source: 'src/formats/pptx/roundtrip/package-preservation.ts',
    tests: ['test/pptx/roundtrip-patch-package.test.ts'],
  },
  {
    estimatedSeconds: 30,
    name: 'patch-error',
    source: 'src/formats/pptx/roundtrip/patch-error.ts',
    tests: ['test/pptx/roundtrip-patch-error.test.ts'],
  },
  {
    estimatedSeconds: 150,
    name: 'relationships',
    source: 'src/formats/pptx/roundtrip/relationships.ts',
    tests: ['test/pptx/roundtrip-relationships.test.ts'],
  },
  {
    estimatedSeconds: 240,
    name: 'shape-range',
    source: 'src/formats/pptx/roundtrip/shape-range.ts',
    tests: ['test/pptx/roundtrip-shape-range.test.ts'],
  },
  {
    estimatedSeconds: 150,
    name: 'text-xml',
    source: 'src/formats/pptx/roundtrip/text-xml.ts',
    tests: ['test/pptx/roundtrip-text-xml.test.ts'],
  },
  {
    estimatedSeconds: 150,
    name: 'transform-xml',
    source: 'src/formats/pptx/roundtrip/transform-xml.ts',
    tests: ['test/pptx/roundtrip-transform-xml.test.ts'],
  },
  {
    estimatedSeconds: 105,
    name: 'writer-shape',
    source: 'src/formats/pptx/writer/shape.ts',
    tests: [
      'test/pptx/writer-shape.test.ts',
      'test/pptx/writer-text-shape.test.ts',
    ],
  },
];

export function resolveMutationModule(name) {
  const result = mutationModules.find((candidate) => candidate.name === name);
  if (result === undefined) {
    throw new Error(
      `Unknown mutation module ${JSON.stringify(name)}; expected one of ${mutationModules
        .map((candidate) => candidate.name)
        .join(', ')}`,
    );
  }
  return result;
}
