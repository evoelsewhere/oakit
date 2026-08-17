import { describe, expect, it } from 'vitest';

import { createPptxWithDependencies } from '../../src/formats/pptx/creator';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import { PptxWriteError } from '../../src/formats/pptx/write-error';

function emptyScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [],
    themes: [],
  };
}

describe('PowerPoint creator error boundaries', () => {
  it('validates before calling package dependencies', async () => {
    const scene = emptyScene();
    scene.size.width = 0;
    let archiveCalls = 0;
    let verifyCalls = 0;

    const promise = createPptxWithDependencies(scene, {
      serializeArchive: () => {
        archiveCalls += 1;
        return Promise.resolve(new Uint8Array());
      },
      verify: () => {
        verifyCalls += 1;
        return Promise.resolve();
      },
    });

    await expect(promise).rejects.toMatchObject({
      code: 'invalid-scene',
      message: 'PowerPoint scene is not valid for creation',
    });
    expect(archiveCalls).toBe(0);
    expect(verifyCalls).toBe(0);
  });

  it('enforces creation-specific validation before building parts', async () => {
    const scene = emptyScene();
    scene.slides = [
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
                    { key: 'run-1', text: 'Missing geometry', type: 'run' },
                  ],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide-1',
      },
    ];
    let archiveCalls = 0;

    const promise = createPptxWithDependencies(scene, {
      serializeArchive: () => {
        archiveCalls += 1;
        return Promise.resolve(new Uint8Array());
      },
      verify: () => Promise.resolve(),
    });

    await expect(promise).rejects.toMatchObject({
      code: 'invalid-scene',
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 requires an authored text transform',
          path: '$.slides[0].elements[0].authored.transform',
        },
      ],
    });
    expect(archiveCalls).toBe(0);
  });

  it('wraps package failures without attempting verification', async () => {
    const cause = new Error('archive failed');
    let verifyCalls = 0;

    const promise = createPptxWithDependencies(emptyScene(), {
      serializeArchive: () => Promise.reject(cause),
      verify: () => {
        verifyCalls += 1;
        return Promise.resolve();
      },
    });

    await expect(promise).rejects.toEqual(
      new PptxWriteError(
        'package-build-failed',
        'Failed to build PowerPoint package',
        { cause },
      ),
    );
    expect(verifyCalls).toBe(0);
  });

  it('wraps strict verification failures and exposes no data', async () => {
    const cause = new Error('verification failed');
    const data = new Uint8Array([1, 2, 3]);

    const promise = createPptxWithDependencies(emptyScene(), {
      serializeArchive: () => Promise.resolve(data),
      verify: () => Promise.reject(cause),
    });

    await expect(promise).rejects.toEqual(
      new PptxWriteError(
        'verification-failed',
        'Generated PowerPoint package failed strict verification',
        { cause },
      ),
    );
    await expect(promise).rejects.not.toHaveProperty('data');
  });

  it('returns the exact C1 report after successful verification', async () => {
    const data = new Uint8Array([4, 5, 6]);
    let verifiedData: Uint8Array | undefined;

    const result = await createPptxWithDependencies(emptyScene(), {
      serializeArchive: () => Promise.resolve(data),
      verify: (value) => {
        verifiedData = value;
        return Promise.resolve();
      },
    });

    expect(verifiedData).toBe(data);
    expect(result.data).toBe(data);
    expect(result.report).toEqual({
      addedPartCount: 9,
      copiedPartCount: 0,
      diagnostics: [],
      level: 'C2',
      operations: [],
      patchedPartCount: 0,
      producerEvidence: [],
      rebuiltPartCount: 0,
      removedPartCount: 0,
      supportProfile: {
        effectiveLevel: 'C2',
        id: 'pptx-create-text-v1',
        producerMatrix: [],
        version: '1',
      },
    });
  });
});
