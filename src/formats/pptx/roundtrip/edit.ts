import { resolvePptxResourceLimits } from '../internal/resource-limits';
import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { isValidXmlText } from '../scene-validation';
import { PptxWriteError } from '../write-error';
import { degreesToAngle, pointsToEmu } from '../writer/units';
import type {
  PptxSceneImageElement,
  PptxSceneShapeElement,
  PptxSceneTableElement,
  PptxSceneTextElement,
  PptxSceneTransform,
} from '../scene-types';
import { canonicalJson } from './canonical-json';
import {
  createPptxRoundTripNativeEditSupportProfile,
  createPptxRoundTripTextEditSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import type {
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetTransformOperation,
  PptxRoundTripSnapshot,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';
import { scalePptxTableIntegerSizes } from './transform-xml';

export interface PptxRoundTripReplaceTextRequest {
  targetKey: string;
  value: string;
}

export interface PptxRoundTripSetTransformRequest {
  targetKey: string;
  value: PptxSceneTransform;
}

function scaledTableSizes(
  values: readonly number[],
  replacementTotal: number,
): number[] {
  const source = values.map(pointsToEmu);
  const replacement = pointsToEmu(replacementTotal);
  return scalePptxTableIntegerSizes(source, replacement).map(
    (value) => value * RATIO_EMUs_Points,
  );
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
          if (
            (element.type === 'image' ||
              element.type === 'shape' ||
              element.type === 'table' ||
              element.type === 'text') &&
            element.key === operation.targetKey
          ) {
            if (element.type === 'table') {
              element.columns = scaledTableSizes(
                element.columns,
                operation.value.width,
              );
              const rowHeights = scaledTableSizes(
                element.rows.map((row) => row.height),
                operation.value.height,
              );
              element.rows.forEach((row, index) => {
                row.height = rowHeights[index] as number;
              });
            }
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

type PptxTransformElement =
  | PptxSceneImageElement
  | PptxSceneShapeElement
  | PptxSceneTableElement
  | PptxSceneTextElement;

function findTransformElement(
  snapshot: PptxRoundTripSnapshot,
  targetKey: string,
  targetType: PptxTransformElement['type'],
): PptxTransformElement {
  let matched: PptxTransformElement | undefined;
  for (const slide of snapshot.document.slides) {
    for (const element of slide.elements) {
      if (element.type !== targetType || element.key !== targetKey) continue;
      matched = element;
    }
  }
  if (matched === undefined) {
    invalidEdit(
      targetType === 'text'
        ? 'PowerPoint transform target key does not exist'
        : `PowerPoint ${targetType} transform target key does not exist`,
    );
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
  try {
    pointsToEmu(value.x);
    pointsToEmu(value.y);
    pointsToEmu(value.width);
    pointsToEmu(value.height);
    degreesToAngle(value.rotation ?? 0);
  } catch {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
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

function validateTableTransformSize(
  target: PptxSceneTableElement,
  value: PptxSceneTransform,
): void {
  if (
    pointsToEmu(value.width) < target.columns.length ||
    pointsToEmu(value.height) < target.rows.length
  ) {
    invalidEdit(
      'PowerPoint table transform is too small for its column and row grid',
    );
  }
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
  snapshot.supportProfile =
    snapshot.supportProfile.id === 'pptx-roundtrip-native-v1'
      ? createPptxRoundTripNativeEditSupportProfile()
      : createPptxRoundTripTextEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}

async function setPptxRoundTripTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
  targetType: PptxTransformElement['type'],
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  const snapshot = structuredClone(validated);
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint transform target key must be a non-empty string');
  }
  const target = findTransformElement(snapshot, request.targetKey, targetType);
  const expectedTransform = target.resolved.transform;
  if (expectedTransform === undefined) {
    invalidEdit('PowerPoint transform target has no resolved transform');
  }
  const transform = normalizePptxRoundTripTransform(request.value);
  if (target.type === 'table') {
    validateTableTransformSize(target, transform);
  }
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
  snapshot.supportProfile =
    targetType === 'image' ||
    targetType === 'shape' ||
    targetType === 'table' ||
    snapshot.supportProfile.id === 'pptx-roundtrip-native-v1'
      ? createPptxRoundTripNativeEditSupportProfile()
      : createPptxRoundTripTextEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}

export function setPptxRoundTripTextTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'text');
}

export function setPptxRoundTripShapeTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'shape');
}

export function setPptxRoundTripImageTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'image');
}

export function setPptxRoundTripTableTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'table');
}
