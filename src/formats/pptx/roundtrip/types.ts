import type { PptxSceneDocument } from '../scene-types';
import type { PptxParseOptions, PptxResourceLimits } from '../types';
import type { PptxWriteSupportProfile } from '../write-types';

export type PptxRoundTripConformance = 'strict' | 'transitional' | 'unknown';

export interface PptxRoundTripRuntimeSource {
  byteLength: number;
  conformance: PptxRoundTripConformance;
  data: Blob | Uint8Array;
  kind: 'bytes';
  sha256: string;
}

export interface PptxRoundTripPortableSource {
  byteLength: number;
  conformance: PptxRoundTripConformance;
  kind: 'base64';
  packageBase64: string;
  sha256: string;
}

export interface PptxSnapshotConsistency {
  canonicalizationVersion: string;
  capabilityProfileVersion: string;
  contractVersion: string;
  hashAlgorithm: 'sha256';
  keyAlgorithmVersion: string;
  operationsSha256: string;
  semanticPreviewSha256: string;
  sourceManifestSha256: string;
}

export interface PptxRoundTripSnapshot {
  consistency: PptxSnapshotConsistency;
  document: PptxSceneDocument;
  format: 'pptx';
  operations: [];
  schemaVersion: 1;
  source: PptxRoundTripRuntimeSource;
  supportProfile: PptxWriteSupportProfile;
}

export interface PptxRoundTripPortableJson {
  consistency: PptxSnapshotConsistency;
  document: PptxSceneDocument;
  format: 'pptx';
  operations: [];
  schemaVersion: 1;
  source: PptxRoundTripPortableSource;
  supportProfile: PptxWriteSupportProfile;
}

export type PptxRoundTripReadOptions = PptxParseOptions;

export interface PptxRoundTripWriteOptions {
  limits?: PptxResourceLimits;
}

export interface PptxRoundTripPortableLimits {
  maxBase64Characters?: number;
  maxDecodedBytes?: number;
}
