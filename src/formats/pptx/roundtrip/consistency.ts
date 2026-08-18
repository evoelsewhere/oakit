import type { PptxSceneDocument, PptxSceneElement } from '../scene-types';
import type { PptxWriteSupportProfile } from '../write-types';
import { canonicalSha256 } from './digest';
import type {
  PptxRoundTripConformance,
  PptxSnapshotConsistency,
} from './types';

export const PPTX_ROUND_TRIP_CONTRACT_VERSION = '1';
export const PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION = 'pptx-scene-key-v1';
export const PPTX_ROUND_TRIP_CANONICALIZATION_VERSION = 'canonical-json-v1';
export const PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION =
  'pptx-roundtrip-text-v1';
export const PPTX_ROUND_TRIP_NATIVE_CAPABILITY_PROFILE_VERSION =
  'pptx-roundtrip-native-v1';

export interface PptxRoundTripConsistencyInput {
  document: PptxSceneDocument;
  operations: readonly unknown[];
  source: {
    byteLength: number;
    conformance: PptxRoundTripConformance;
    sha256: string;
  };
  supportProfile: PptxWriteSupportProfile;
}

function elementKeyManifest(element: PptxSceneElement): unknown {
  return element.type === 'text'
    ? {
        key: element.key,
        paragraphs: element.text.paragraphs.map((paragraph) => ({
          children: paragraph.children.map((child) => child.key),
          key: paragraph.key,
        })),
      }
    : { key: element.key };
}

function keyManifest(document: PptxSceneDocument): unknown {
  return {
    layouts: document.layouts.map((layout) => ({
      elements: layout.elements.map(elementKeyManifest),
      key: layout.key,
    })),
    masters: document.masters.map((master) => ({
      elements: master.elements.map(elementKeyManifest),
      key: master.key,
    })),
    slides: document.slides.map((slide) => ({
      elements: slide.elements.map(elementKeyManifest),
      key: slide.key,
    })),
    themes: document.themes.map((theme) => theme.key),
  };
}

export function createPptxRoundTripSupportProfile(): PptxWriteSupportProfile {
  return {
    effectiveLevel: 'R0',
    id: 'pptx-roundtrip-r0',
    producerMatrix: [],
    version: '1',
  };
}

export function createPptxRoundTripTextEditSupportProfile(): PptxWriteSupportProfile {
  return {
    effectiveLevel: 'R2',
    id: 'pptx-roundtrip-text-v1',
    producerMatrix: [],
    version: '1',
  };
}

export function createPptxRoundTripNativeEditSupportProfile(): PptxWriteSupportProfile {
  return {
    effectiveLevel: 'R2',
    id: 'pptx-roundtrip-native-v1',
    producerMatrix: [],
    version: '1',
  };
}

export async function createPptxSnapshotConsistency(
  input: PptxRoundTripConsistencyInput,
): Promise<PptxSnapshotConsistency> {
  const [sourceManifestSha256, semanticPreviewSha256, operationsSha256] =
    await Promise.all([
      canonicalSha256({
        format: 'pptx',
        keyManifest: keyManifest(input.document),
        schemaVersion: 1,
        source: {
          byteLength: input.source.byteLength,
          conformance: input.source.conformance,
          sha256: input.source.sha256,
        },
        supportProfile: input.supportProfile,
      }),
      canonicalSha256(input.document),
      canonicalSha256(input.operations),
    ]);

  return {
    canonicalizationVersion: PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
    capabilityProfileVersion:
      input.supportProfile.id === 'pptx-roundtrip-native-v1'
        ? PPTX_ROUND_TRIP_NATIVE_CAPABILITY_PROFILE_VERSION
        : PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
    contractVersion: PPTX_ROUND_TRIP_CONTRACT_VERSION,
    hashAlgorithm: 'sha256',
    keyAlgorithmVersion: PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
    operationsSha256,
    semanticPreviewSha256,
    sourceManifestSha256,
  };
}
