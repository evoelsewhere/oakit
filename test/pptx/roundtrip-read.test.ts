import { describe, expect, it } from 'vitest';

import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import { PptxWriteError } from '../../src/formats/pptx/write-error';
import { assertValidPptxRoundTripPreview } from '../../src/formats/pptx/roundtrip/read';

function invalidPreview(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 0 },
    slides: [],
    themes: [],
  };
}

describe('PowerPoint round-trip reader invariants', () => {
  it('rejects an internally produced preview that violates the scene contract', () => {
    let received: unknown;
    try {
      assertValidPptxRoundTripPreview(invalidPreview());
    } catch (error) {
      received = error;
    }

    expect(received).toBeInstanceOf(PptxWriteError);
    expect(received).toMatchObject({
      code: 'invalid-snapshot',
      issues: [
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.size.width',
        },
      ],
      message: 'PowerPoint semantic preview is not valid for round-trip',
    });
  });
});
