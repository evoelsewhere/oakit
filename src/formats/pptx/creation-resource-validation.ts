import {
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
} from './creation-limits';
import type {
  PptxSceneSlide,
  PptxSceneTextElement,
  PptxSceneValidationIssue,
} from './scene-types';

type JsonObject = Record<string, unknown>;

function stringCodeUnits(value: unknown): number {
  if (typeof value === 'string') return value.length;
  if (value === null) return 0;
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
  paragraphs: number;
  textNodes: number;
}

function countCreationResources(document: JsonObject): CreationResourceCounts {
  const counts: CreationResourceCounts = {
    elements: 0,
    paragraphs: 0,
    textNodes: 0,
  };
  const slides = document.slides as PptxSceneSlide[];
  slides.forEach((slide) => {
    const elements = slide.elements;
    counts.elements += elements.length;
    elements.forEach((element) => {
      const textElement = element as PptxSceneTextElement;
      const paragraphs = textElement.text.paragraphs;
      counts.paragraphs += paragraphs.length;
      paragraphs.forEach((paragraph) => {
        counts.textNodes += paragraph.children.length;
      });
    });
  });
  return counts;
}

export function validatePowerPointCreationResources(
  document: JsonObject,
): PptxSceneValidationIssue[] {
  const issues: PptxSceneValidationIssue[] = [];
  const counts = countCreationResources(document);
  if (counts.elements > MAX_POWERPOINT_CREATION_ELEMENTS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_ELEMENTS} elements`,
      path: '$.slides',
    });
  }
  if (counts.paragraphs > MAX_POWERPOINT_CREATION_PARAGRAPHS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_PARAGRAPHS} paragraphs`,
      path: '$.slides',
    });
  }
  if (counts.textNodes > MAX_POWERPOINT_CREATION_TEXT_NODES) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_TEXT_NODES} text nodes`,
      path: '$.slides',
    });
  }
  const codeUnits = stringCodeUnits(document);
  if (codeUnits > MAX_POWERPOINT_CREATION_STRING_CODE_UNITS) {
    issues.push({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_STRING_CODE_UNITS} string code units`,
      path: '$',
    });
  }
  return issues;
}
