import type { PptxSceneRunProperties, PptxSceneTextNode } from '../scene-types';
import { serializeSolidColorFill } from './color';
import { pointsToFontSize } from './units';
import { escapeXmlAttribute, serializeDrawingText } from './xml';

export interface PptxTextSerializationContext {
  allocateFieldId(): string;
}

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function serializeProperties(
  tagName: 'a:endParaRPr' | 'a:rPr',
  properties: PptxSceneRunProperties | undefined,
): string {
  const attributes: string[] = [];
  const children: string[] = [];
  if (properties?.language !== undefined) {
    attributes.push(`lang="${escapeXmlAttribute(properties.language)}"`);
  }
  if (properties?.fontSize !== undefined) {
    attributes.push(`sz="${pointsToFontSize(properties.fontSize)}"`);
  }
  if (properties?.bold !== undefined) {
    attributes.push(`b="${booleanAttribute(properties.bold)}"`);
  }
  if (properties?.italic !== undefined) {
    attributes.push(`i="${booleanAttribute(properties.italic)}"`);
  }
  if (properties?.color !== undefined) {
    children.push(serializeSolidColorFill(properties.color));
  }
  const attributeText = attributes.length > 0 ? ` ${attributes.join(' ')}` : '';
  if (properties?.fontFamily !== undefined) {
    const typeface = escapeXmlAttribute(properties.fontFamily);
    children.push(
      `<a:latin typeface="${typeface}"/><a:ea typeface="${typeface}"/><a:cs typeface="${typeface}"/>`,
    );
  }
  if (children.length === 0) {
    return `<${tagName}${attributeText}/>`;
  }
  return `<${tagName}${attributeText}>${children.join('')}</${tagName}>`;
}

export function serializeRunProperties(
  properties?: PptxSceneRunProperties,
): string {
  return serializeProperties('a:rPr', properties);
}

export function serializeEndParagraphProperties(
  properties?: PptxSceneRunProperties,
): string {
  return serializeProperties('a:endParaRPr', properties);
}

export function serializeTextNode(
  node: PptxSceneTextNode,
  context: PptxTextSerializationContext,
): string {
  const properties = serializeRunProperties(node.properties);
  switch (node.type) {
    case 'run':
      return `<a:r>${properties}${serializeDrawingText(node.text, node.preserveSpace)}</a:r>`;
    case 'field':
      return `<a:fld id="${escapeXmlAttribute(context.allocateFieldId())}" type="${escapeXmlAttribute(node.fieldType)}">${properties}${serializeDrawingText(node.text)}</a:fld>`;
    case 'break':
      return `<a:br>${properties}</a:br>`;
  }
}
