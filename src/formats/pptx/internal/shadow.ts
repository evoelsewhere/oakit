import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';
import type { Shadow } from '../types';

import { getSolidFill } from './fill';
import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { getTextByPathList } from '../../../common';

export function getShadow(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): Shadow {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj);
  const outerShdwAttrs =
    getTextByPathList<Record<string, string>>(node, ['attrs']) ?? {};
  const dir = outerShdwAttrs['dir']
    ? parseInt(outerShdwAttrs['dir']) / 60000
    : 0;
  const dist = outerShdwAttrs['dist']
    ? parseInt(outerShdwAttrs['dist']) * RATIO_EMUs_Points
    : 0;
  const blurRad = outerShdwAttrs['blurRad']
    ? parseInt(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points
    : 0;
  const vx = dist * Math.sin((dir * Math.PI) / 180);
  const hx = dist * Math.cos((dir * Math.PI) / 180);

  return {
    h: hx,
    v: vx,
    blur: blurRad,
    color: chdwClrNode,
  };
}
