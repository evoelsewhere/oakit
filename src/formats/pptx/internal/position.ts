import type { XmlLookupValue } from '../../../common';

import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { getTextByPathList, numberToFixed } from '../../../common';

interface Position {
  left: number;
  top: number;
}

interface Size {
  height: number;
  width: number;
}

function emuToPoints(value: unknown): number {
  const emus = Number(value ?? 0);
  return numberToFixed(emus * RATIO_EMUs_Points);
}

export function getPosition(
  slideSpNode?: XmlLookupValue,
  slideLayoutSpNode?: XmlLookupValue,
  slideMasterSpNode?: XmlLookupValue,
): Position {
  const off =
    getTextByPathList(slideSpNode, ['a:off', 'attrs']) ||
    getTextByPathList(slideLayoutSpNode, ['a:off', 'attrs']) ||
    getTextByPathList(slideMasterSpNode, ['a:off', 'attrs']);

  if (!off) return { top: 0, left: 0 };

  return {
    top: emuToPoints(off['y']),
    left: emuToPoints(off['x']),
  };
}

export function getSize(
  slideSpNode?: XmlLookupValue,
  slideLayoutSpNode?: XmlLookupValue,
  slideMasterSpNode?: XmlLookupValue,
): Size {
  const ext =
    getTextByPathList(slideSpNode, ['a:ext', 'attrs']) ||
    getTextByPathList(slideLayoutSpNode, ['a:ext', 'attrs']) ||
    getTextByPathList(slideMasterSpNode, ['a:ext', 'attrs']);

  if (!ext) return { width: 0, height: 0 };

  return {
    width: emuToPoints(ext['cx']),
    height: emuToPoints(ext['cy']),
  };
}
