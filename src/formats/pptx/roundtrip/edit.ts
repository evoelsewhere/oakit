import { resolvePptxResourceLimits } from '../internal/resource-limits';
import { isValidXmlText } from '../scene-validation';
import { PptxWriteError } from '../write-error';
import {
  createPptxRoundTripTextEditSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import type {
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSnapshot,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';

export interface PptxRoundTripReplaceTextRequest {
  targetKey: string;
  value: string;
}

export function applyPptxRoundTripOperationsToPreview(
  snapshot: PptxRoundTripSnapshot,
): PptxRoundTripSnapshot['document'] {
  const document = structuredClone(snapshot.document);
  for (const operation of snapshot.operations) {
    let applied = false;
    for (const slide of document.slides) {
      for (const element of slide.elements) {
        if (element.type !== 'text') continue;
        for (const paragraph of element.text.paragraphs) {
          for (const child of paragraph.children) {
            if (child.type === 'run' && child.key === operation.targetKey) {
              child.text = operation.value;
              applied = true;
            }
          }
        }
      }
    }
    if (!applied) {
      throw new PptxWriteError(
        'verification-failed',
        `PowerPoint text edit verification target disappeared: ${operation.targetKey}`,
      );
    }
  }
  return document;
}

function invalidEdit(message: string): never {
  throw new PptxWriteError('invalid-edit-operation', message);
}

function findRunText(
  snapshot: PptxRoundTripSnapshot,
  targetKey: string,
): string {
  let matched: string | undefined;
  for (const slide of snapshot.document.slides) {
    for (const element of slide.elements) {
      if (element.type !== 'text') continue;
      for (const paragraph of element.text.paragraphs) {
        for (const child of paragraph.children) {
          if (child.key !== targetKey || child.type !== 'run') continue;
          if (matched !== undefined) {
            invalidEdit('PowerPoint text edit target key is ambiguous');
          }
          matched = child.text;
        }
      }
    }
  }
  if (matched === undefined) {
    invalidEdit('PowerPoint text edit target key does not exist');
  }
  return matched;
}

function validateRequest(
  request: PptxRoundTripReplaceTextRequest,
  maxXmlBytes: number,
): void {
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint text edit target key must be a non-empty string');
  }
  if (typeof request.value !== 'string') {
    invalidEdit('PowerPoint text edit value must be a string');
  }
  if (!isValidXmlText(request.value)) {
    invalidEdit('PowerPoint text edit value is not safe XML text');
  }
  if (
    request.value.length > maxXmlBytes ||
    new TextEncoder().encode(request.value).byteLength > maxXmlBytes
  ) {
    invalidEdit('PowerPoint text edit value exceeds the XML part byte limit');
  }
}

export async function replacePptxRoundTripText(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripReplaceTextRequest,
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  validateRequest(request, limits.maxXmlBytes);
  const snapshot = structuredClone(validated);
  const expectedText = findRunText(snapshot, request.targetKey);
  if (expectedText === request.value) {
    invalidEdit('PowerPoint text edit must change the target value');
  }
  if (
    snapshot.operations.some(
      (operation) => operation.targetKey === request.targetKey,
    )
  ) {
    invalidEdit('PowerPoint text edit target is already scheduled');
  }

  const operation: PptxRoundTripReplaceTextOperation = {
    expectedText,
    id: `replace-text-${snapshot.operations.length + 1}`,
    kind: 'replace-text',
    targetKey: request.targetKey,
    value: request.value,
  };
  snapshot.operations.push(operation);
  snapshot.supportProfile = createPptxRoundTripTextEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}
