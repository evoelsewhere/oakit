import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import type { PptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/types';
import { validatePptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/validate';

const HASH = 'a'.repeat(64);

function snapshot(): PptxRoundTripSnapshot {
  return {
    consistency: {
      canonicalizationVersion: 'canonical-json-v1',
      capabilityProfileVersion: 'pptx-roundtrip-r0-v1',
      contractVersion: '1',
      hashAlgorithm: 'sha256',
      keyAlgorithmVersion: 'pptx-scene-key-v1',
      operationsSha256: HASH,
      semanticPreviewSha256: HASH,
      sourceManifestSha256: HASH,
    },
    document: {
      layouts: [],
      masters: [],
      media: [],
      schemaVersion: 2,
      size: { height: 540, width: 960 },
      slides: [],
      themes: [],
    },
    format: 'pptx',
    operations: [],
    schemaVersion: 1,
    source: {
      byteLength: 1,
      conformance: 'unknown',
      data: new Uint8Array([1]),
      kind: 'bytes',
      sha256: HASH,
    },
    supportProfile: {
      effectiveLevel: 'R0',
      id: 'pptx-roundtrip-r0',
      producerMatrix: [],
      version: '1',
    },
  };
}

function rootRecord(value: PptxRoundTripSnapshot): Record<string, unknown> {
  return value as unknown as Record<string, unknown>;
}

function nestedRecord(
  value: PptxRoundTripSnapshot,
  key: 'consistency' | 'source' | 'supportProfile',
): Record<string, unknown> {
  return value[key] as unknown as Record<string, unknown>;
}

function expectInvalid(value: unknown, code: string, message: string): void {
  let received: unknown;
  try {
    validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits());
  } catch (error) {
    received = error;
  }
  expect(received).toMatchObject({ code, message });
}

describe('PowerPoint round-trip snapshot contract validation', () => {
  it('accepts the exact R0 runtime shape without replacing it', () => {
    const value = snapshot();
    expect(
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
    ).toBe(value);
  });

  it.each([
    [
      'primitive root',
      () => null,
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'extra root field',
      () => {
        const value = snapshot();
        rootRecord(value).extra = true;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'missing root field',
      () => {
        const value = snapshot();
        delete rootRecord(value).document;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'substituted root field',
      () => {
        const value = snapshot();
        delete rootRecord(value).document;
        rootRecord(value).extra = true;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'schema version',
      () => {
        const value = snapshot();
        rootRecord(value).schemaVersion = 2;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot schema version is unsupported',
    ],
    [
      'format',
      () => {
        const value = snapshot();
        rootRecord(value).format = 'docx';
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot format must be pptx',
    ],
    [
      'operations type',
      () => {
        const value = snapshot();
        rootRecord(value).operations = {};
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot operations must be an array',
    ],
    [
      'nonempty operations',
      () => {
        const value = snapshot();
        rootRecord(value).operations = [{ type: 'replaceText' }];
        return value;
      },
      'unsupported-edit-operation',
      'PowerPoint R0 round-trip does not support edit operations',
    ],
  ])('rejects %s', (_name, create, code, message) => {
    expectInvalid(create(), code, message);
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).source = {};
      },
      'PowerPoint round-trip snapshot source has an invalid shape',
    ],
    [
      'extra field',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').extra = true;
      },
      'PowerPoint round-trip snapshot source has an invalid shape',
    ],
    [
      'kind',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').kind = 'base64';
      },
      'PowerPoint round-trip snapshot source kind must be bytes',
    ],
    [
      'data',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').data = {};
      },
      'PowerPoint round-trip snapshot source data must be Uint8Array or Blob',
    ],
    [
      'zero length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 0;
        nestedRecord(value, 'source').data = new Uint8Array();
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'fractional length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 1.5;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'negative length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = -1;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'unsafe length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = Number.MAX_SAFE_INTEGER + 1;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'declared length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 2;
      },
      'PowerPoint round-trip snapshot source byteLength does not match its data',
    ],
    [
      'conformance',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').conformance = 'future';
      },
      'PowerPoint round-trip snapshot source conformance is invalid',
    ],
    [
      'SHA-256',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = 'A'.repeat(64);
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
    [
      'SHA-256 prefix',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = `x${HASH}`;
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
    [
      'SHA-256 suffix',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = `${HASH}x`;
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
  ])('rejects source %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('accepts every declared source conformance and Blob transport', () => {
    for (const conformance of ['strict', 'transitional', 'unknown'] as const) {
      const value = snapshot();
      value.source.conformance = conformance;
      value.source.data = new Blob([new Uint8Array([1]).buffer]);
      expect(() =>
        validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
      ).not.toThrow();
    }
  });

  it('applies the runtime input byte limit to snapshot source data', () => {
    expect(() =>
      validatePptxRoundTripSnapshot(
        snapshot(),
        resolvePptxResourceLimits({ maxInputBytes: 1 }),
      ),
    ).not.toThrow();
    const value = snapshot();
    value.source.byteLength = 2;
    value.source.data = new Uint8Array([1, 2]);
    expect(() =>
      validatePptxRoundTripSnapshot(
        value,
        resolvePptxResourceLimits({ maxInputBytes: 1 }),
      ),
    ).toThrow('maxInputBytes exceeded');
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).supportProfile = {};
      },
      'PowerPoint round-trip snapshot support profile has an invalid shape',
    ],
    [
      'level',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').effectiveLevel = 'R1';
      },
      'PowerPoint round-trip snapshot support level must be R0',
    ],
    [
      'id',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').id = 'other';
      },
      'PowerPoint round-trip snapshot support profile id is unsupported',
    ],
    [
      'version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').version = '2';
      },
      'PowerPoint round-trip snapshot support profile version is unsupported',
    ],
    [
      'producer matrix type',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').producerMatrix = {};
      },
      'PowerPoint round-trip snapshot producer matrix must be an array',
    ],
    [
      'producer claim',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').producerMatrix = ['PowerPoint'];
      },
      'PowerPoint round-trip R0 snapshot cannot claim producer evidence',
    ],
  ])('rejects support profile %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).consistency = {};
      },
      'PowerPoint round-trip snapshot consistency has an invalid shape',
    ],
    [
      'canonicalization version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').canonicalizationVersion = 'future';
      },
      'PowerPoint round-trip snapshot canonicalization version is unsupported',
    ],
    [
      'capability profile version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').capabilityProfileVersion = 'future';
      },
      'PowerPoint round-trip snapshot capability profile version is unsupported',
    ],
    [
      'contract version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').contractVersion = '2';
      },
      'PowerPoint round-trip snapshot contract version is unsupported',
    ],
    [
      'hash algorithm',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').hashAlgorithm = 'sha512';
      },
      'PowerPoint round-trip snapshot hash algorithm is unsupported',
    ],
    [
      'key algorithm version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').keyAlgorithmVersion = 'future';
      },
      'PowerPoint round-trip snapshot key algorithm version is unsupported',
    ],
    [
      'operations digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').operationsSha256 = 'short';
      },
      'PowerPoint round-trip snapshot operations SHA-256 is invalid',
    ],
    [
      'preview digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').semanticPreviewSha256 = 'short';
      },
      'PowerPoint round-trip snapshot semantic preview SHA-256 is invalid',
    ],
    [
      'source digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').sourceManifestSha256 = 'short';
      },
      'PowerPoint round-trip snapshot source manifest SHA-256 is invalid',
    ],
  ])('rejects consistency %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('rejects an invalid semantic preview with its scene issues', () => {
    const value = snapshot();
    value.document.size.width = 0;

    let received: unknown;
    try {
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits());
    } catch (error) {
      received = error;
    }
    expect(received).toMatchObject({
      code: 'invalid-snapshot',
      issues: [
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.size.width',
        },
      ],
      message: 'PowerPoint round-trip snapshot semantic preview is invalid',
    });
  });
});
