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
  let v = getTextByPathList<string | number>(slideNode, [
    'p:txBody',
    'a:bodyPr',
    'attrs',
    attrName,
  ]);
  if (v !== undefined && v !== null && v !== '') return v;

  v = getTextByPathList<string | number>(layoutNode, [
    'p:txBody',
    'a:bodyPr',
    'attrs',
    attrName,
  ]);
  if (v !== undefined && v !== null && v !== '') return v;

  return getTextByPathList<string | number>(masterNode, [
    'p:txBody',
    'a:bodyPr',
    'attrs',
    attrName,
  ]);
}

function emuToPt(emuStr: string | number | null | undefined): number | null {
  if (emuStr === undefined || emuStr === null || emuStr === '') return null;
  const v = parseInt(String(emuStr), 10);
  if (!Number.isFinite(v)) return null;
  return numberToFixed(v * RATIO_EMUs_Points);
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

  if (!nodeBodyPr) {
    if (!layoutBodyPr) {
      if (!masterBodyPr) return null;
    }
  }

  let li = getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'lIns');
  if (li === undefined || li === null || li === '') li = DEFAULT_INSET_EMU.lIns;

  let ti = getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'tIns');
  if (ti === undefined || ti === null || ti === '') ti = DEFAULT_INSET_EMU.tIns;

  let ri = getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'rIns');
  if (ri === undefined || ri === null || ri === '') ri = DEFAULT_INSET_EMU.rIns;

  let bi = getInsetAttr(node, slideLayoutSpNode, slideMasterSpNode, 'bIns');
  if (bi === undefined || bi === null || bi === '') bi = DEFAULT_INSET_EMU.bIns;

  let l = emuToPt(li);
  if (l === null) l = 0;

  let t = emuToPt(ti);
  if (t === null) t = 0;

  let r = emuToPt(ri);
  if (r === null) r = 0;

  let b = emuToPt(bi);
  if (b === null) b = 0;

  return { l, t, r, b };
}
