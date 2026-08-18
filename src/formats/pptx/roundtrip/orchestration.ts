import {
  assertPptxArchiveWithinLimits,
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import type { PptxDocument } from '../types';
import {
  assertSafeEditablePptxPackage,
  decodeEditablePptxXml,
  generatePptxPatchedArchive,
  readPptxPartPayloads,
  verifyPptxPatchedPayloads,
} from './package-preservation';
import { unsupportedPptxEdit } from './patch-error';
import { resolvePptxSlideParts } from './relationships';
import { patchPptxShapeTextXml } from './text-xml';
import { patchPptxShapeTransformXml } from './transform-xml';
import type {
  PptxRoundTripOperation,
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetTransformOperation,
} from './types';
import JSZip from 'jszip';

const TARGET_KEY_PATTERN = /^slide-([1-9]\d*)-element-([1-9]\d*)-run-1$/;
const TRANSFORM_TARGET_KEY_PATTERN = /^slide-([1-9]\d*)-element-([1-9]\d*)$/;

interface TextTarget {
  elementIndex: number;
  shapeId: string;
  slideIndex: number;
}

export interface PptxPatchedPackage {
  copiedPartCount: number;
  data: Uint8Array;
  patchedPartCount: number;
}

function textTarget(
  operation: PptxRoundTripReplaceTextOperation,
  document: PptxDocument,
): TextTarget {
  const match = TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedPptxEdit(
      'PowerPoint text edit target is not a supported slide text run key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndex = Number(match[2]) - 1;
  if (
    !Number.isSafeInteger(slideIndex) ||
    !Number.isSafeInteger(elementIndex)
  ) {
    unsupportedPptxEdit('PowerPoint text edit target index is unsafe');
  }
  const element = document.slides[slideIndex]?.elements[elementIndex];
  if (element?.type !== 'text') {
    unsupportedPptxEdit(
      'PowerPoint text edit target is not a slide-owned text element',
    );
  }
  return { elementIndex, shapeId: element.id, slideIndex };
}

function transformTarget(
  operation: PptxRoundTripSetTransformOperation,
  document: PptxDocument,
): TextTarget {
  const match = TRANSFORM_TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedPptxEdit(
      'PowerPoint transform target is not a supported slide text element key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndex = Number(match[2]) - 1;
  if (
    !Number.isSafeInteger(slideIndex) ||
    !Number.isSafeInteger(elementIndex)
  ) {
    unsupportedPptxEdit('PowerPoint transform target index is unsafe');
  }
  const element = document.slides[slideIndex]?.elements[elementIndex];
  if (element?.type !== 'shape' && element?.type !== 'text') {
    unsupportedPptxEdit(
      'PowerPoint transform target is not a slide-owned text or shape element',
    );
  }
  return { elementIndex, shapeId: element.id, slideIndex };
}

export async function patchPptxOperations(
  bytes: Uint8Array,
  document: PptxDocument,
  operations: readonly PptxRoundTripOperation[],
  limits: ResolvedPptxResourceLimits,
): Promise<PptxPatchedPackage> {
  const archive = await JSZip.loadAsync(bytes);
  assertPptxArchiveWithinLimits(archive, limits);
  assertSafeEditablePptxPackage(archive);
  const [slides, sourcePayloads] = await Promise.all([
    resolvePptxSlideParts(archive, limits),
    readPptxPartPayloads(archive, limits),
  ]);
  if (slides.length !== document.slides.length) {
    unsupportedPptxEdit(
      'PowerPoint text edit slide order does not match the parsed document',
    );
  }

  const patchedParts = new Set<string>();
  const editedXml = new Map<string, string>();
  for (const operation of operations) {
    const target =
      operation.kind === 'replace-text'
        ? textTarget(operation, document)
        : transformTarget(operation, document);
    const slidePart = slides[target.slideIndex] as string;
    const sourceBytes = sourcePayloads.get(slidePart) as Uint8Array;
    const current =
      editedXml.get(slidePart) ?? decodeEditablePptxXml(sourceBytes, limits);
    const patched =
      operation.kind === 'replace-text'
        ? patchPptxShapeTextXml(current, target.shapeId, operation)
        : patchPptxShapeTransformXml(current, target.shapeId, operation);
    editedXml.set(slidePart, patched);
    patchedParts.add(slidePart);
  }

  for (const [part, xml] of editedXml) {
    const entry = archive.file(part) as JSZip.JSZipObject;
    archive.file(part, xml, {
      date: entry.date,
    });
  }
  const output = await generatePptxPatchedArchive(archive);
  assertPptxInputWithinLimits(output, limits);

  const outputArchive = await JSZip.loadAsync(output);
  assertPptxArchiveWithinLimits(outputArchive, limits);
  const outputPayloads = await readPptxPartPayloads(outputArchive, limits);
  verifyPptxPatchedPayloads(sourcePayloads, outputPayloads, patchedParts);

  return {
    copiedPartCount: sourcePayloads.size - patchedParts.size,
    data: output,
    patchedPartCount: patchedParts.size,
  };
}
