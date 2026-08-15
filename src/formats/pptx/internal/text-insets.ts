import type { XmlLookupValue } from '../../../common';

import { getTextByPathList, numberToFixed } from '../../../common';
import { RATIO_EMUs_Points } from '../../../common/ooxml/units';

const DEFAULT_INSET_EMU = {
  lIns: 91440, // 0.1 in
  rIns: 91440, // 0.1 in
  tIns: 45720, // 0.05 in
  bIns: 45720, // 0.05 in
};

function getInsetAttr(
  slideNode: XmlLookupValue | undefined,
  layoutNode: XmlLookupValue | undefined,
  masterNode: XmlLookupValue | undefined,
  attrName: string,
): string | number | undefined {
  for (const candidate of [slideNode, layoutNode, masterNode]) {
    const value = getTextByPathList<string | number>(candidate, [
      'p:txBody',
      'a:bodyPr',
      'attrs',
      attrName,
    ]);
    if (value !== undefined && value !== '') return value;
  }
  return undefined;
}

function emuToPt(value: string | number): number {
  const emus = Number(value);
  return Number.isFinite(emus) ? numberToFixed(emus * RATIO_EMUs_Points) : 0;
}

export function getTextInsets(
  node: XmlLookupValue,
  slideLayoutSpNode?: XmlLookupValue,
  slideMasterSpNode?: XmlLookupValue,
): { b: number; l: number; r: number; t: number } | null {
  const nodeBodyPr = getTextByPathList(node, ['p:txBody', 'a:bodyPr']);
  const layoutBodyPr = getTextByPathList(slideLayoutSpNode, [
    'p:txBody',
    'a:bodyPr',
  ]);
  const masterBodyPr = getTextByPathList(slideMasterSpNode, [
    'p:txBody',
    'a:bodyPr',
  ]);

  if (!nodeBodyPr && !layoutBodyPr && !masterBodyPr) return null;

  const l = emuToPt(
    getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'lIns') ??
      DEFAULT_INSET_EMU.lIns,
  );
  const t = emuToPt(
    getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'tIns') ??
      DEFAULT_INSET_EMU.tIns,
  );
  const r = emuToPt(
    getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'rIns') ??
      DEFAULT_INSET_EMU.rIns,
  );
  const b = emuToPt(
    getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'bIns') ??
      DEFAULT_INSET_EMU.bIns,
  );

  return { b, l, r, t };
}
