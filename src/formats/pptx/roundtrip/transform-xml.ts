import { decodeXmlEntities } from '../../../common/text/html';
import { degreesToAngle, pointsToEmu } from '../writer/units';
import { unsupportedPptxEdit } from './patch-error';
import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditableGraphicFrameXml,
  resolvePptxEditableGroupXml,
  resolvePptxEditablePictureXml,
  resolvePptxEditableShapeXml,
} from './shape-range';
import type { PptxRoundTripSetTransformOperation } from './types';
import type { PptxSceneGroupTransform } from '../scene-types';

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

function patchPptxTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
  resolveElement: typeof resolvePptxEditableShapeXml,
  transformNamespace: 'drawing' | 'presentation' = 'drawing',
): string {
  const { drawingPrefix, markupPrefix, presentationPrefix, range, shape } =
    resolveElement(xml, shapeId);
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

  const transformName = qualifiedPptxName(
    transformNamespace === 'drawing' ? drawingPrefix : presentationPrefix,
    'xfrm',
  );
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

function replaceIntegerAttribute(
  tag: string,
  name: string,
  value: number,
): string {
  const pattern = new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`);
  return tag.replace(pattern, `$1"${value}"`);
}

export function scalePptxTableIntegerSizes(
  source: readonly number[],
  replacementTotal: number,
): number[] {
  const sourceTotal = source.reduce((total, value) => total + value, 0);
  const replacements: number[] = [];
  let allocated = 0;
  source.forEach((sourceValue, index) => {
    const remaining = source.length - index - 1;
    const value =
      index === source.length - 1
        ? replacementTotal - allocated
        : Math.max(
            1,
            Math.min(
              Math.round((sourceValue * replacementTotal) / sourceTotal),
              replacementTotal - allocated - remaining,
            ),
          );
    replacements.push(value);
    allocated += value;
  });
  return replacements;
}

function scaleTableAttributeTags(
  shape: string,
  tagName: string,
  attributeName: string,
  expectedTotal: number,
  replacementTotal: number,
  description: string,
): string {
  const attributePattern =
    '((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|\'[^\']*\'))*)';
  const tagPattern = new RegExp(
    `<${escapePptxXmlPattern(tagName)}${attributePattern}\\s*\\/?>`,
    'g',
  );
  const matches = [...shape.matchAll(tagPattern)];
  if (matches.length === 0) {
    unsupportedPptxEdit(`PowerPoint table has no ${description}`);
  }
  const source = matches.map((match) =>
    integerAttribute(match[1] as string, attributeName),
  );
  if (
    source.some((value) => value <= 0) ||
    source.reduce((total, value) => total + value, 0) !== expectedTotal
  ) {
    unsupportedPptxEdit(
      `PowerPoint table ${description} do not match the preview precondition`,
    );
  }
  if (replacementTotal < matches.length) {
    unsupportedPptxEdit(
      `PowerPoint table ${description} cannot fit the requested transform`,
    );
  }
  const replacements = scalePptxTableIntegerSizes(source, replacementTotal);

  let result = shape;
  for (let index = matches.length - 1; index >= 0; index -= 1) {
    const match = matches[index] as RegExpMatchArray;
    const start = match.index as number;
    const replacement = replaceIntegerAttribute(
      match[0],
      attributeName,
      replacements[index] as number,
    );
    result = `${result.slice(0, start)}${replacement}${result.slice(start + match[0].length)}`;
  }
  return result;
}

function scalePptxTableGridXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  const { drawingPrefix, range, shape } = resolvePptxEditableGraphicFrameXml(
    xml,
    shapeId,
  );
  const tableName = qualifiedPptxName(drawingPrefix, 'tbl');
  if (!pptxShapeHasElement(shape, tableName)) {
    unsupportedPptxEdit('PowerPoint graphic frame is not a native table');
  }
  const columns = scaleTableAttributeTags(
    shape,
    qualifiedPptxName(drawingPrefix, 'gridCol'),
    'w',
    pointsToEmu(operation.expectedTransform.width),
    pointsToEmu(operation.value.width),
    'column widths',
  );
  const rows = scaleTableAttributeTags(
    columns,
    qualifiedPptxName(drawingPrefix, 'tr'),
    'h',
    pointsToEmu(operation.expectedTransform.height),
    pointsToEmu(operation.value.height),
    'row heights',
  );
  return `${xml.slice(0, range.start)}${rows}${xml.slice(range.end)}`;
}

export function patchPptxShapeTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  return patchPptxTransformXml(
    xml,
    shapeId,
    operation,
    resolvePptxEditableShapeXml,
  );
}

export function patchPptxPictureTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  return patchPptxTransformXml(
    xml,
    shapeId,
    operation,
    resolvePptxEditablePictureXml,
  );
}

export function patchPptxGraphicFrameTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  const transformed = patchPptxTransformXml(
    xml,
    shapeId,
    operation,
    resolvePptxEditableGraphicFrameXml,
    'presentation',
  );
  return scalePptxTableGridXml(transformed, shapeId, operation);
}

export function patchPptxChartFrameTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  return patchPptxTransformXml(
    xml,
    shapeId,
    operation,
    resolvePptxEditableGraphicFrameXml,
    'presentation',
  );
}

function requiredGroupTransform(
  value: PptxRoundTripSetTransformOperation['value'],
): PptxSceneGroupTransform {
  if (!('childSpace' in value)) {
    unsupportedPptxEdit('PowerPoint group transform has no child space');
  }
  return value;
}

export function patchPptxGroupTransformXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
  const { drawingPrefix, markupPrefix, presentationPrefix, range, shape } =
    resolvePptxEditableGroupXml(xml, shapeId);
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
      'PowerPoint group transform target contains unsupported compatibility markup',
    );
  }
  const transformName = qualifiedPptxName(drawingPrefix, 'xfrm');
  const groupPropertiesName = qualifiedPptxName(presentationPrefix, 'grpSpPr');
  const offsetName = qualifiedPptxName(drawingPrefix, 'off');
  const extentName = qualifiedPptxName(drawingPrefix, 'ext');
  const childOffsetName = qualifiedPptxName(drawingPrefix, 'chOff');
  const childExtentName = qualifiedPptxName(drawingPrefix, 'chExt');
  const attributePattern =
    '((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|\'[^\']*\'))*)';
  const transformPattern = new RegExp(
    `<${escapePptxXmlPattern(transformName)}${attributePattern}\\s*>\\s*` +
      `<${escapePptxXmlPattern(offsetName)}${attributePattern}\\s*\\/>\\s*` +
      `<${escapePptxXmlPattern(extentName)}${attributePattern}\\s*\\/>\\s*` +
      `<${escapePptxXmlPattern(childOffsetName)}${attributePattern}\\s*\\/>\\s*` +
      `<${escapePptxXmlPattern(childExtentName)}${attributePattern}\\s*\\/>\\s*` +
      `<\\/${escapePptxXmlPattern(transformName)}\\s*>`,
    'g',
  );
  const groupPropertiesPattern = new RegExp(
    `<${escapePptxXmlPattern(groupPropertiesName)}${attributePattern}\\s*>([\\s\\S]*?)<\\/${escapePptxXmlPattern(groupPropertiesName)}\\s*>`,
    'g',
  );
  const groupPropertiesMatches = [...shape.matchAll(groupPropertiesPattern)];
  if (groupPropertiesMatches.length === 0) {
    unsupportedPptxEdit(
      'PowerPoint group target must contain one direct group property block',
    );
  }
  const groupPropertiesMatch = groupPropertiesMatches[0] as RegExpMatchArray;
  const groupProperties = groupPropertiesMatch[2] as string;
  const matches = [...groupProperties.matchAll(transformPattern)];
  if (matches.length !== 1) {
    unsupportedPptxEdit(
      'PowerPoint group target must contain one simple group transform',
    );
  }
  const match = matches[0] as RegExpMatchArray;
  const source = {
    childHeight: integerAttribute(match[5] as string, 'cy'),
    childWidth: integerAttribute(match[5] as string, 'cx'),
    childX: integerAttribute(match[4] as string, 'x'),
    childY: integerAttribute(match[4] as string, 'y'),
    flipHorizontal: booleanAttribute(match[1] as string, 'flipH'),
    flipVertical: booleanAttribute(match[1] as string, 'flipV'),
    height: integerAttribute(match[3] as string, 'cy'),
    rotation: optionalIntegerAttribute(match[1] as string, 'rot'),
    width: integerAttribute(match[3] as string, 'cx'),
    x: integerAttribute(match[2] as string, 'x'),
    y: integerAttribute(match[2] as string, 'y'),
  };
  const expected = requiredGroupTransform(operation.expectedTransform);
  const expectedXml = {
    childHeight: pointsToEmu(expected.childSpace.height),
    childWidth: pointsToEmu(expected.childSpace.width),
    childX: pointsToEmu(expected.childSpace.x),
    childY: pointsToEmu(expected.childSpace.y),
    flipHorizontal: expected.flipHorizontal ?? false,
    flipVertical: expected.flipVertical ?? false,
    height: pointsToEmu(expected.height),
    rotation: degreesToAngle(expected.rotation ?? 0),
    width: pointsToEmu(expected.width),
    x: pointsToEmu(expected.x),
    y: pointsToEmu(expected.y),
  };
  if (JSON.stringify(source) !== JSON.stringify(expectedXml)) {
    unsupportedPptxEdit(
      'PowerPoint group source XML does not match its preview precondition',
    );
  }
  const replacement = requiredGroupTransform(operation.value);
  const replacementAttributes = [
    replacement.rotation === 0
      ? ''
      : ` rot="${degreesToAngle(replacement.rotation ?? 0)}"`,
    replacement.flipHorizontal ? ' flipH="1"' : '',
    replacement.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const child = replacement.childSpace;
  const replacementXml =
    `<${transformName}${replacementAttributes}>` +
    `<${offsetName} x="${pointsToEmu(replacement.x)}" y="${pointsToEmu(replacement.y)}"/>` +
    `<${extentName} cx="${pointsToEmu(replacement.width)}" cy="${pointsToEmu(replacement.height)}"/>` +
    `<${childOffsetName} x="${pointsToEmu(child.x)}" y="${pointsToEmu(child.y)}"/>` +
    `<${childExtentName} cx="${pointsToEmu(child.width)}" cy="${pointsToEmu(child.height)}"/>` +
    `</${transformName}>`;
  const propertiesStart =
    (groupPropertiesMatch.index ?? 0) +
    groupPropertiesMatch[0].indexOf(groupProperties);
  const matchStart = range.start + propertiesStart + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacementXml}${xml.slice(matchEnd)}`;
}
