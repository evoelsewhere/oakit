import type { PptxSceneTextElement, PptxSceneTransform } from '../scene-types';
import { serializeTextBody } from './text-body';
import type { PptxTextSerializationContext } from './text-node';
import { degreesToAngle, pointsToEmu } from './units';
import { escapeXmlAttribute } from './xml';

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

export function serializeShapeTransform(transform: PptxSceneTransform): string {
  const attributes: string[] = [];
  if (transform.rotation !== undefined) {
    attributes.push(`rot="${degreesToAngle(transform.rotation)}"`);
  }
  if (transform.flipHorizontal !== undefined) {
    attributes.push(`flipH="${booleanAttribute(transform.flipHorizontal)}"`);
  }
  if (transform.flipVertical !== undefined) {
    attributes.push(`flipV="${booleanAttribute(transform.flipVertical)}"`);
  }
  const attributeText = attributes.length > 0 ? ` ${attributes.join(' ')}` : '';
  return `<a:xfrm${attributeText}><a:off x="${pointsToEmu(transform.x)}" y="${pointsToEmu(transform.y)}"/><a:ext cx="${pointsToEmu(transform.width)}" cy="${pointsToEmu(transform.height)}"/></a:xfrm>`;
}

function serializeNonVisualProperties(
  element: PptxSceneTextElement,
  shapeId: number,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `Text Box ${shapeId}`)}"`,
  ];
  if (element.description !== undefined) {
    attributes.push(`descr="${escapeXmlAttribute(element.description)}"`);
  }
  if (element.title !== undefined) {
    attributes.push(`title="${escapeXmlAttribute(element.title)}"`);
  }
  if (element.authored.hidden !== undefined) {
    attributes.push(`hidden="${booleanAttribute(element.authored.hidden)}"`);
  }
  return `<p:nvSpPr><p:cNvPr ${attributes.join(' ')}/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`;
}

export function serializeTextShape(
  element: PptxSceneTextElement,
  transform: PptxSceneTransform,
  shapeId: number,
  context: PptxTextSerializationContext,
): string {
  const nonVisual = serializeNonVisualProperties(element, shapeId);
  const shapeProperties = `<p:spPr>${serializeShapeTransform(transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>`;
  return `<p:sp>${nonVisual}${shapeProperties}${serializeTextBody(element.text, context)}</p:sp>`;
}
