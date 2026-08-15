import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';
import type { Shadow } from '../types';

import { getSolidFill } from './fill';
import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { getTextByPathList } from '../../../common';

function nonNegativeInteger(value: string | undefined): number {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? Math.max(parsed, 0) : 0;
}

export function getShadow(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): Shadow {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj);
  const outerShdwAttrs =
    getTextByPathList<Record<string, string>>(node, ['attrs']) ?? {};
  const dir = nonNegativeInteger(outerShdwAttrs['dir']) / 60_000;
  const dist = nonNegativeInteger(outerShdwAttrs['dist']) * RATIO_EMUs_Points;
  const blurRad =
    nonNegativeInteger(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points;
  const vx = dist * Math.sin((dir * Math.PI) / 180);
  const hx = dist * Math.cos((dir * Math.PI) / 180);

  return {
    h: hx,
    v: vx,
    blur: blurRad,
    color: chdwClrNode,
  };
}
