import { createHash } from 'node:crypto';

import { describe, expect, it } from 'vitest';

import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import {
  createPptxRoundTripSupportProfile,
  createPptxSnapshotConsistency,
  PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
  PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
  PPTX_ROUND_TRIP_CONTRACT_VERSION,
  PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
} from '../../src/formats/pptx/roundtrip/consistency';

function sha256(value: string): string {
  return createHash('sha256').update(value, 'utf8').digest('hex');
}

function scene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {},
            key: 'text-1',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'run-1', text: 'hello', type: 'run' },
                    { key: 'break-1', type: 'break' },
                  ],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
          {
            authored: {},
            feature: 'chart',
            key: 'unsupported-1',
            resolved: { hidden: false },
            type: 'unsupported',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

function hierarchyScene(): PptxSceneDocument {
  const unsupported = (key: string) => ({
    authored: {},
    feature: 'shape',
    key,
    resolved: { hidden: false },
    type: 'unsupported' as const,
  });
  return {
    layouts: [
      {
        elements: [unsupported('layout-element')],
        key: 'layout-1',
        masterKey: 'master-1',
      },
    ],
    masters: [
      {
        elements: [unsupported('master-element')],
        key: 'master-1',
        themeKey: 'theme-1',
      },
    ],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [],
    themes: [{ key: 'theme-1' }],
  };
}

function tableScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {},
            columns: [100],
            key: 'table-1',
            resolved: { hidden: false },
            rows: [
              {
                cells: [
                  {
                    text: {
                      body: {},
                      paragraphs: [
                        {
                          children: [
                            {
                              key: 'table-run',
                              text: 'value',
                              type: 'run',
                            },
                          ],
                          key: 'table-paragraph',
                        },
                      ],
                    },
                  },
                ],
                height: 40,
              },
            ],
            type: 'table',
          },
        ],
        key: 'table-slide',
      },
    ],
    themes: [],
  };
}

describe('PowerPoint round-trip snapshot consistency', () => {
  it('locks the R0 support and version identifiers', () => {
    expect(createPptxRoundTripSupportProfile()).toEqual({
      effectiveLevel: 'R0',
      id: 'pptx-roundtrip-r0',
      producerMatrix: [],
      version: '1',
    });
    expect({
      canonicalizationVersion: PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
      capabilityProfileVersion: PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
      contractVersion: PPTX_ROUND_TRIP_CONTRACT_VERSION,
      keyAlgorithmVersion: PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
    }).toEqual({
      canonicalizationVersion: 'canonical-json-v1',
      capabilityProfileVersion: 'pptx-roundtrip-text-v1',
      contractVersion: '1',
      keyAlgorithmVersion: 'pptx-scene-key-v1',
    });
  });

  it('hashes source metadata, stable keys, preview, and operations independently', async () => {
    const document = scene();
    const supportProfile = createPptxRoundTripSupportProfile();
    const consistency = await createPptxSnapshotConsistency({
      document,
      operations: [],
      source: {
        byteLength: 3,
        conformance: 'strict',
        sha256: 'abc',
      },
      supportProfile,
    });

    const expectedManifest =
      '{"format":"pptx","keyManifest":{"layouts":[],"masters":[],"slides":[{"elements":[{"key":"text-1","paragraphs":[{"children":["run-1","break-1"],"key":"paragraph-1"}]},{"key":"unsupported-1"}],"key":"slide-1"}],"themes":[]},"schemaVersion":1,"source":{"byteLength":3,"conformance":"strict","sha256":"abc"},"supportProfile":{"effectiveLevel":"R0","id":"pptx-roundtrip-r0","producerMatrix":[],"version":"1"}}';
    const expectedPreview =
      '{"layouts":[],"masters":[],"media":[],"schemaVersion":2,"size":{"height":540,"width":960},"slides":[{"elements":[{"authored":{},"key":"text-1","resolved":{"hidden":false},"text":{"body":{},"paragraphs":[{"children":[{"key":"run-1","text":"hello","type":"run"},{"key":"break-1","type":"break"}],"key":"paragraph-1"}]},"type":"text"},{"authored":{},"feature":"chart","key":"unsupported-1","resolved":{"hidden":false},"type":"unsupported"}],"key":"slide-1"}],"themes":[]}';

    expect(consistency).toEqual({
      canonicalizationVersion: 'canonical-json-v1',
      capabilityProfileVersion: 'pptx-roundtrip-text-v1',
      contractVersion: '1',
      hashAlgorithm: 'sha256',
      keyAlgorithmVersion: 'pptx-scene-key-v1',
      operationsSha256: sha256('[]'),
      semanticPreviewSha256: sha256(expectedPreview),
      sourceManifestSha256: sha256(expectedManifest),
    });
  });

  it('binds key topology separately from semantic preview content', async () => {
    const original = scene();
    const changedKey = scene();
    const changedText = scene();
    const keySlide = changedKey.slides[0];
    if (keySlide === undefined) throw new Error('Expected key slide');
    const keyElement = keySlide.elements[0];
    if (keyElement === undefined) throw new Error('Expected key element');
    keyElement.key = 'text-2';
    const textSlide = changedText.slides[0];
    if (textSlide === undefined) throw new Error('Expected text slide');
    const textElement = textSlide.elements[0];
    if (textElement === undefined) throw new Error('Expected text element');
    if (textElement.type !== 'text') throw new Error('Expected text element');
    const paragraph = textElement.text.paragraphs[0];
    if (paragraph === undefined) throw new Error('Expected paragraph');
    const run = paragraph.children[0];
    if (run === undefined) throw new Error('Expected text run');
    if (run.type !== 'run') throw new Error('Expected text run');
    run.text = 'changed';
    const source = {
      byteLength: 3,
      conformance: 'strict' as const,
      sha256: 'abc',
    };
    const supportProfile = createPptxRoundTripSupportProfile();

    const consistencyFor = (document: PptxSceneDocument) =>
      createPptxSnapshotConsistency({
        document,
        operations: [],
        source,
        supportProfile,
      });
    const [baseline, keyResult, textResult] = await Promise.all([
      consistencyFor(original),
      consistencyFor(changedKey),
      consistencyFor(changedText),
    ]);

    expect(keyResult.sourceManifestSha256).not.toBe(
      baseline.sourceManifestSha256,
    );
    expect(keyResult.semanticPreviewSha256).not.toBe(
      baseline.semanticPreviewSha256,
    );
    expect(textResult.sourceManifestSha256).toBe(baseline.sourceManifestSha256);
    expect(textResult.semanticPreviewSha256).not.toBe(
      baseline.semanticPreviewSha256,
    );
  });

  it('preserves operation order in its own digest', async () => {
    const document = scene();
    const source = {
      byteLength: 3,
      conformance: 'strict' as const,
      sha256: 'abc',
    };
    const supportProfile = createPptxRoundTripSupportProfile();

    const forward = await createPptxSnapshotConsistency({
      document,
      operations: [{ id: 'first' }, { id: 'second' }],
      source,
      supportProfile,
    });
    const reverse = await createPptxSnapshotConsistency({
      document,
      operations: [{ id: 'second' }, { id: 'first' }],
      source,
      supportProfile,
    });

    expect(forward.operationsSha256).toBe(
      sha256('[{"id":"first"},{"id":"second"}]'),
    );
    expect(reverse.operationsSha256).not.toBe(forward.operationsSha256);
    expect(reverse.sourceManifestSha256).toBe(forward.sourceManifestSha256);
    expect(reverse.semanticPreviewSha256).toBe(forward.semanticPreviewSha256);
  });

  it('binds hierarchy and nested owner keys into the key manifest', async () => {
    const source = {
      byteLength: 3,
      conformance: 'strict' as const,
      sha256: 'abc',
    };
    const supportProfile = createPptxRoundTripSupportProfile();
    const consistencyFor = (document: PptxSceneDocument) =>
      createPptxSnapshotConsistency({
        document,
        operations: [],
        source,
        supportProfile,
      });
    const baseline = await consistencyFor(hierarchyScene());
    const mutations: Array<(document: PptxSceneDocument) => void> = [
      (document) => {
        const theme = document.themes[0];
        if (theme === undefined) throw new Error('Expected theme');
        theme.key = 'theme-2';
      },
      (document) => {
        const master = document.masters[0];
        if (master === undefined) throw new Error('Expected master');
        master.key = 'master-2';
      },
      (document) => {
        const master = document.masters[0];
        if (master === undefined) throw new Error('Expected master');
        const element = master.elements[0];
        if (element === undefined) throw new Error('Expected master element');
        element.key = 'master-element-2';
      },
      (document) => {
        const layout = document.layouts[0];
        if (layout === undefined) throw new Error('Expected layout');
        layout.key = 'layout-2';
      },
      (document) => {
        const layout = document.layouts[0];
        if (layout === undefined) throw new Error('Expected layout');
        const element = layout.elements[0];
        if (element === undefined) throw new Error('Expected layout element');
        element.key = 'layout-element-2';
      },
    ];

    for (const mutate of mutations) {
      const changed = hierarchyScene();
      mutate(changed);
      const result = await consistencyFor(changed);
      expect(result.sourceManifestSha256).not.toBe(
        baseline.sourceManifestSha256,
      );
    }
  });

  it('binds every native table cell text key into the source manifest', async () => {
    const supportProfile = createPptxRoundTripSupportProfile();
    const consistency = await createPptxSnapshotConsistency({
      document: tableScene(),
      operations: [],
      source: { byteLength: 3, conformance: 'strict', sha256: 'abc' },
      supportProfile,
    });
    const expectedManifest =
      '{"format":"pptx","keyManifest":{"layouts":[],"masters":[],"slides":[{"elements":[{"key":"table-1","rows":[[[{"children":["table-run"],"key":"table-paragraph"}]]]}],"key":"table-slide"}],"themes":[]},"schemaVersion":1,"source":{"byteLength":3,"conformance":"strict","sha256":"abc"},"supportProfile":{"effectiveLevel":"R0","id":"pptx-roundtrip-r0","producerMatrix":[],"version":"1"}}';

    expect(consistency.sourceManifestSha256).toBe(sha256(expectedManifest));
  });

  it('returns independent mutable support profile values', () => {
    const first = createPptxRoundTripSupportProfile();
    first.producerMatrix.push('caller-change');

    expect(createPptxRoundTripSupportProfile().producerMatrix).toEqual([]);
  });
});
