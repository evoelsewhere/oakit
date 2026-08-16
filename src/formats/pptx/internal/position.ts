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

function emuToPoints(value: unknown, allowNegative: boolean): number {
  if (value === undefined || value === null) return 0;
  if (typeof value !== 'string' && typeof value !== 'number') {
    return Number.NaN;
  }
  const lexicalValue = typeof value === 'string' ? value : value.toString(10);
  if (!/^[+-]?(?:0|[1-9]\d*)$/.test(lexicalValue)) return Number.NaN;
  const emus = Number(lexicalValue);
  if (!Number.isSafeInteger(emus) || (!allowNegative && emus < 0)) {
    return Number.NaN;
  }
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
    top: emuToPoints(off['y'], true),
    left: emuToPoints(off['x'], true),
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
    width: emuToPoints(ext['cx'], false),
    height: emuToPoints(ext['cy'], false),
  };
}
