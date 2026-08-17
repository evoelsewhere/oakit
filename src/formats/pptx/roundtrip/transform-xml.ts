import { decodeXmlEntities } from '../../../common/text/html';
import { degreesToAngle, pointsToEmu } from '../writer/units';
import { unsupportedPptxEdit } from './patch-error';
import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditableShapeXml,
} from './shape-range';
import type { PptxRoundTripSetTransformOperation } from './types';

function attributeValue(attributesText: string, name: string): string | null {
  const pattern = new RegExp(
    `(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
  );
  const match = pattern.exec(attributesText);
  return match === null
    ? null
    : decodeXmlEntities((match[1] ?? match[2]) as string);
}

function integerAttribute(attributesText: string, name: string): number {
  const value = attributeValue(attributesText, name);
  if (value === null || !/^-?\d+$/.test(value)) {
    unsupportedPptxEdit(`PowerPoint transform ${name} attribute is invalid`);
  }
  const result = Number(value);
  if (!Number.isSafeInteger(result)) {
    unsupportedPptxEdit(`PowerPoint transform ${name} attribute is unsafe`);
  }
  return result;
}

function optionalIntegerAttribute(
  attributesText: string,
  name: string,
): number {
  return attributeValue(attributesText, name) === null
    ? 0
    : integerAttribute(attributesText, name);
}

function booleanAttribute(attributesText: string, name: string): boolean {
  const value = attributeValue(attributesText, name);
  if (value === null || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  unsupportedPptxEdit(`PowerPoint transform ${name} attribute is invalid`);
}

export function patchPptxShapeTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  const { drawingPrefix, markupPrefix, presentationPrefix, range, shape } =
    resolvePptxEditableShapeXml(xml, shapeId);
  if (
    (markupPrefix !== undefined &&
      pptxShapeHasElement(
        shape,
        qualifiedPptxName(markupPrefix, 'AlternateContent'),
      )) ||
    pptxShapeHasElement(
      shape,
      qualifiedPptxName(presentationPrefix, 'extLst'),
    ) ||
    pptxShapeHasElement(shape, qualifiedPptxName(drawingPrefix, 'extLst'))
  ) {
    unsupportedPptxEdit(
      'PowerPoint transform target contains unsupported compatibility markup',
    );
  }

  const transformName = qualifiedPptxName(drawingPrefix, 'xfrm');
  const offsetName = qualifiedPptxName(drawingPrefix, 'off');
  const extentName = qualifiedPptxName(drawingPrefix, 'ext');
  const attributePattern =
    '((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|\'[^\']*\'))*)';
  const transformPattern = new RegExp(
    `<${escapePptxXmlPattern(transformName)}${attributePattern}\\s*>\\s*` +
      `<${escapePptxXmlPattern(offsetName)}${attributePattern}\\s*\\/>\\s*` +
      `<${escapePptxXmlPattern(extentName)}${attributePattern}\\s*\\/>\\s*` +
      `<\\/${escapePptxXmlPattern(transformName)}\\s*>`,
    'g',
  );
  const matches = [...shape.matchAll(transformPattern)];
  if (matches.length !== 1) {
    unsupportedPptxEdit(
      'PowerPoint transform target must contain one simple shape transform',
    );
  }
  const match = matches[0] as RegExpMatchArray;
  const transformAttributes = match[1] as string;
  const offsetAttributes = match[2] as string;
  const extentAttributes = match[3] as string;
  const source = {
    flipHorizontal: booleanAttribute(transformAttributes, 'flipH'),
    flipVertical: booleanAttribute(transformAttributes, 'flipV'),
    height: integerAttribute(extentAttributes, 'cy'),
    rotation: optionalIntegerAttribute(transformAttributes, 'rot'),
    width: integerAttribute(extentAttributes, 'cx'),
    x: integerAttribute(offsetAttributes, 'x'),
    y: integerAttribute(offsetAttributes, 'y'),
  };
  const expected = {
    flipHorizontal: operation.expectedTransform.flipHorizontal as boolean,
    flipVertical: operation.expectedTransform.flipVertical as boolean,
    height: pointsToEmu(operation.expectedTransform.height),
    rotation: degreesToAngle(operation.expectedTransform.rotation ?? 0),
    width: pointsToEmu(operation.expectedTransform.width),
    x: pointsToEmu(operation.expectedTransform.x),
    y: pointsToEmu(operation.expectedTransform.y),
  };
  if (JSON.stringify(source) !== JSON.stringify(expected)) {
    unsupportedPptxEdit(
      'PowerPoint transform source XML does not match its preview precondition',
    );
  }
  const replacementAttributes = [
    operation.value.rotation === 0
      ? ''
      : ` rot="${degreesToAngle(operation.value.rotation ?? 0)}"`,
    operation.value.flipHorizontal ? ' flipH="1"' : '',
    operation.value.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const replacement =
    `<${transformName}${replacementAttributes}>` +
    `<${offsetName} x="${pointsToEmu(operation.value.x)}" y="${pointsToEmu(operation.value.y)}"/>` +
    `<${extentName} cx="${pointsToEmu(operation.value.width)}" cy="${pointsToEmu(operation.value.height)}"/>` +
    `</${transformName}>`;
  const matchStart = range.start + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacement}${xml.slice(matchEnd)}`;
}
