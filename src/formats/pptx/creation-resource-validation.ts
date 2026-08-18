import {
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_MEDIA,
  MAX_POWERPOINT_CREATION_MEDIA_BYTES,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
  MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES,
} from './creation-limits';
import type {
  PptxSceneMedia,
  PptxSceneSlide,
  PptxSceneValidationIssue,
} from './scene-types';

type JsonObject = Record<string, unknown>;

function stringCodeUnits(value: unknown): number {
  if (typeof value === 'string') return value.length;
  if (value === null) return 0;
  if (ArrayBuffer.isView(value)) return 0;
  switch (typeof value) {
    case 'object':
      return Object.values(value).reduce<number>(
        (total, child) => total + stringCodeUnits(child),
        0,
      );
    default:
      return 0;
  }
}

interface CreationResourceCounts {
  elements: number;
  media: number;
  mediaBytes: number;
  maxMediaBytes: number;
  paragraphs: number;
  textNodes: number;
}

function countCreationResources(document: JsonObject): CreationResourceCounts {
  const counts: CreationResourceCounts = {
    elements: 0,
    media: 0,
    mediaBytes: 0,
    maxMediaBytes: 0,
    paragraphs: 0,
    textNodes: 0,
  };
  const slides = document.slides as PptxSceneSlide[];
  slides.forEach((slide) => {
    const elements = slide.elements;
    counts.elements += elements.length;
    elements.forEach((element) => {
      if (element.type !== 'text') return;
      const paragraphs = element.text.paragraphs;
      counts.paragraphs += paragraphs.length;
      paragraphs.forEach((paragraph) => {
        counts.textNodes += paragraph.children.length;
      });
    });
  });
  const media = Array.isArray(document.media)
    ? (document.media as PptxSceneMedia[])
    : [];
  counts.media = media.length;
  for (const item of media) {
    if (item === undefined) continue;
    counts.mediaBytes += item.data.byteLength;
    counts.maxMediaBytes = Math.max(counts.maxMediaBytes, item.data.byteLength);
  }
  return counts;
}

export function validatePowerPointCreationResources(
  document: JsonObject,
  profile: 'create-native-v1' | 'create-text-v1' = 'create-text-v1',
): PptxSceneValidationIssue[] {
  const issues: PptxSceneValidationIssue[] = [];
  const counts = countCreationResources(document);
  if (counts.elements > MAX_POWERPOINT_CREATION_ELEMENTS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_ELEMENTS} elements`,
      path: '$.slides',
    });
  }
  if (counts.media > MAX_POWERPOINT_CREATION_MEDIA) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_MEDIA} media resources`,
      path: '$.media',
    });
  }
  if (counts.maxMediaBytes > MAX_POWERPOINT_CREATION_MEDIA_BYTES) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_MEDIA_BYTES} bytes per media resource`,
      path: '$.media',
    });
  }
  if (counts.mediaBytes > MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES} total media bytes`,
      path: '$.media',
    });
  }
  if (counts.paragraphs > MAX_POWERPOINT_CREATION_PARAGRAPHS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_PARAGRAPHS} paragraphs`,
      path: '$.slides',
    });
  }
  if (counts.textNodes > MAX_POWERPOINT_CREATION_TEXT_NODES) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_TEXT_NODES} text nodes`,
      path: '$.slides',
    });
  }
  const codeUnits = stringCodeUnits(document);
  if (codeUnits > MAX_POWERPOINT_CREATION_STRING_CODE_UNITS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_STRING_CODE_UNITS} string code units`,
      path: '$',
    });
  }
  return issues;
}
