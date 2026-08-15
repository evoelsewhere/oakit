import type { XmlLookupValue } from '../../../common';

import { getTextByPathList } from '../../../common';

export function findTransitionNode(
  content: XmlLookupValue | null | undefined,
  rootElement: string,
): XmlLookupValue | null {
  const path1 = [rootElement, 'p:transition'];
  let transitionNode = getTextByPathList<XmlLookupValue>(content, path1);
  if (transitionNode) return transitionNode;

  const path2 = [
    rootElement,
    'mc:AlternateContent',
    'mc:Choice',
    'p:transition',
  ];
  transitionNode = getTextByPathList<XmlLookupValue>(content, path2);
  if (transitionNode) return transitionNode;

  const path3 = [
    rootElement,
    'mc:AlternateContent',
    'mc:Fallback',
    'p:transition',
  ];
  transitionNode = getTextByPathList<XmlLookupValue>(content, path3);

  return transitionNode ?? null;
}

interface ParsedTransition {
  autoNextAfter?: number;
  direction: string | null;
  duration: number;
  type: string;
}

const DECIMAL_DIGITS = new Set('0123456789');

function parseMilliseconds(value: string | undefined): number | undefined {
  if (
    value === undefined ||
    value.length === 0 ||
    Array.from(value).some((character) => !DECIMAL_DIGITS.has(character))
  ) {
    return undefined;
  }
  const milliseconds = Number(value);
  return Number.isSafeInteger(milliseconds) ? milliseconds : undefined;
}

export function parseTransition(
  transitionNode: XmlLookupValue | null | undefined,
): ParsedTransition | null {
  if (!transitionNode) return null;

  const transition: ParsedTransition = {
    type: 'none',
    duration: 1000,
    direction: null,
  };

  const attrs =
    getTextByPathList<Record<string, string>>(transitionNode, ['attrs']) ?? {};

  let durationFound = false;
  const durRegex = /^p\d{2}:dur$/;
  for (const key in attrs) {
    const duration = parseMilliseconds(attrs[key]);
    if (durRegex.test(key) && duration !== undefined) {
      transition.duration = duration;
      durationFound = true;
      break;
    }
  }

  if (!durationFound && attrs.spd) {
    switch (attrs.spd) {
      case 'med':
        transition.duration = 800;
        break;
      case 'fast':
        transition.duration = 500;
        break;
    }
  }

  const autoNextAfter = parseMilliseconds(attrs.advTm);
  if (attrs.advClick === '0' && autoNextAfter !== undefined) {
    transition.autoNextAfter = autoNextAfter;
  }

  const effectRegex = /^(p|p\d{2}):/;
  for (const key of Object.keys(transitionNode)) {
    if (effectRegex.test(key)) {
      const effectNode = transitionNode[key];
      transition.type = key.substring(key.indexOf(':') + 1);

      const effectAttrs =
        getTextByPathList<Record<string, string>>(effectNode, ['attrs']) ?? {};

      const effectDuration = parseMilliseconds(effectAttrs.dur);
      if (effectDuration !== undefined) {
        if (!durationFound) transition.duration = effectDuration;
      }
      if (effectAttrs.dir) transition.direction = effectAttrs.dir;
      break;
    }
  }

  return transition;
}
