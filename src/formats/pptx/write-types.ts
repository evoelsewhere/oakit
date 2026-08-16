import type { PptxSceneValidationIssue } from './scene-types';

export type PptxFidelityLevel = 'C1' | 'C2' | 'R0' | 'R1' | 'R2' | 'R3';

export type PptxWriteErrorCode =
  'invalid-scene' | 'package-build-failed' | 'verification-failed';

export interface PptxWriteErrorOptions extends ErrorOptions {
  issues?: readonly PptxSceneValidationIssue[];
}

export interface PptxWriteDiagnostic {
  code: string;
  message: string;
  path?: string;
}

export interface PptxOperationEvidence {
  id: string;
  kind: string;
  status: 'applied' | 'verified';
}

export interface PptxWriteSupportProfile {
  effectiveLevel: PptxFidelityLevel;
  id: string;
  producerMatrix: string[];
  version: string;
}

export interface PptxWriteReport {
  addedPartCount: number;
  copiedPartCount: number;
  diagnostics: PptxWriteDiagnostic[];
  level: PptxFidelityLevel;
  operations: PptxOperationEvidence[];
  patchedPartCount: number;
  producerEvidence: string[];
  rebuiltPartCount: number;
  removedPartCount: number;
  supportProfile: PptxWriteSupportProfile;
}

export interface PptxWriteResult {
  data: Uint8Array;
  report: PptxWriteReport;
}
