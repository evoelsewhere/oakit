import { resolvePptxResourceLimits } from '../internal/resource-limits';
import { isValidXmlText } from '../scene-validation';
import { PptxWriteError } from '../write-error';
import type { PptxSceneTextElement, PptxSceneTransform } from '../scene-types';
import { canonicalJson } from './canonical-json';
import {
  createPptxRoundTripTextEditSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import type {
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetTransformOperation,
  PptxRoundTripSnapshot,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';

export interface PptxRoundTripReplaceTextRequest {
  targetKey: string;
  value: string;
}

export interface PptxRoundTripSetTransformRequest {
  targetKey: string;
  value: PptxSceneTransform;
}

export function applyPptxRoundTripOperationsToPreview(
  snapshot: PptxRoundTripSnapshot,
): PptxRoundTripSnapshot['document'] {
  const document = structuredClone(snapshot.document);
  for (const operation of snapshot.operations) {
    if (operation.kind === 'set-transform') {
      let applied = false;
      for (const slide of document.slides) {
        for (const element of slide.elements) {
          if (element.type === 'text' && element.key === operation.targetKey) {
            element.resolved.transform = structuredClone(operation.value);
            applied = true;
          }
        }
      }
      if (!applied) {
        throw new PptxWriteError(
          'verification-failed',
          `PowerPoint transform verification target disappeared: ${operation.targetKey}`,
        );
      }
      continue;
    }
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

function findTextElement(
  snapshot: PptxRoundTripSnapshot,
  targetKey: string,
): PptxSceneTextElement {
  let matched: PptxSceneTextElement | undefined;
  for (const slide of snapshot.document.slides) {
    for (const element of slide.elements) {
      if (element.type !== 'text' || element.key !== targetKey) continue;
      matched = element;
    }
  }
  if (matched === undefined) {
    invalidEdit('PowerPoint transform target key does not exist');
  }
  return matched;
}

export function normalizePptxRoundTripTransform(
  value: PptxSceneTransform,
): PptxSceneTransform {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    invalidEdit('PowerPoint transform value must be an object');
  }
  const allowedKeys = new Set([
    'flipHorizontal',
    'flipVertical',
    'height',
    'rotation',
    'width',
    'x',
    'y',
  ]);
  if (Object.keys(value).some((key) => !allowedKeys.has(key))) {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  for (const key of ['x', 'y', 'width', 'height'] as const) {
    if (!Number.isFinite(value[key])) {
      invalidEdit('PowerPoint transform value is not a valid scene transform');
    }
  }
  if (value.width <= 0 || value.height <= 0) {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  if (value.rotation !== undefined && !Number.isFinite(value.rotation)) {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  for (const key of ['flipHorizontal', 'flipVertical'] as const) {
    if (value[key] !== undefined && typeof value[key] !== 'boolean') {
      invalidEdit('PowerPoint transform value is not a valid scene transform');
    }
  }
  return {
    flipHorizontal: value.flipHorizontal ?? false,
    flipVertical: value.flipVertical ?? false,
    height: value.height,
    rotation: value.rotation ?? 0,
    width: value.width,
    x: value.x,
    y: value.y,
  };
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

export function validatePptxRoundTripReplaceTextRequest(
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
  if (new TextEncoder().encode(request.value).byteLength > maxXmlBytes) {
    invalidEdit('PowerPoint text edit value exceeds the XML part byte limit');
  }
}

export async function replacePptxRoundTripText(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripReplaceTextRequest,
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  validatePptxRoundTripReplaceTextRequest(request, limits.maxXmlBytes);
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

export async function setPptxRoundTripTextTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  const snapshot = structuredClone(validated);
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint transform target key must be a non-empty string');
  }
  const target = findTextElement(snapshot, request.targetKey);
  const expectedTransform = target.resolved.transform;
  if (expectedTransform === undefined) {
    invalidEdit('PowerPoint transform target has no resolved transform');
  }
  const transform = normalizePptxRoundTripTransform(request.value);
  if (canonicalJson(expectedTransform) === canonicalJson(transform)) {
    invalidEdit('PowerPoint transform edit must change the target value');
  }
  if (
    snapshot.operations.some(
      (operation) => operation.targetKey === request.targetKey,
    )
  ) {
    invalidEdit('PowerPoint transform target is already scheduled');
  }
  const operation: PptxRoundTripSetTransformOperation = {
    expectedTransform: structuredClone(expectedTransform),
    id: `set-transform-${snapshot.operations.length + 1}`,
    kind: 'set-transform',
    targetKey: request.targetKey,
    value: transform,
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
