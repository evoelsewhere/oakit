import type {
  PptxSceneElementBase,
  PptxSceneShapeElement,
  PptxSceneTransform,
} from '../scene-types';
import { serializeSolidColorFill } from './color';
import { degreesToAngle, pointsToEmu } from './units';
import { escapeXmlAttribute } from './xml';

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function serializeTransform(
  transform: PptxSceneTransform,
  tagName: 'a:xfrm' | 'p:xfrm',
): string {
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
  return `<${tagName}${attributeText}><a:off x="${pointsToEmu(transform.x)}" y="${pointsToEmu(transform.y)}"/><a:ext cx="${pointsToEmu(transform.width)}" cy="${pointsToEmu(transform.height)}"/></${tagName}>`;
}

export function serializeShapeTransform(transform: PptxSceneTransform): string {
  return serializeTransform(transform, 'a:xfrm');
}

export function serializeGraphicFrameTransform(
  transform: PptxSceneTransform,
): string {
  return serializeTransform(transform, 'p:xfrm');
}

export function serializeShapeNonVisualProperties(
  element: PptxSceneElementBase,
  shapeId: number,
  textBox: boolean,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `${textBox ? 'Text Box' : 'Shape'} ${shapeId}`)}"`,
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
  return `<p:nvSpPr><p:cNvPr ${attributes.join(' ')}/><p:cNvSpPr${textBox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>`;
}

export function serializeNativeShapeProperties(
  element: PptxSceneElementBase,
  transform: PptxSceneTransform,
): string {
  const authored = element.authored;
  const geometry = authored.geometry ?? 'rect';
  const fill =
    authored.fillColor === undefined
      ? '<a:noFill/>'
      : serializeSolidColorFill(authored.fillColor);
  const line =
    authored.lineColor === undefined && authored.lineWidth === undefined
      ? '<a:ln><a:noFill/></a:ln>'
      : `<a:ln${authored.lineWidth === undefined ? '' : ` w="${pointsToEmu(authored.lineWidth)}"`}>${serializeSolidColorFill(authored.lineColor ?? '#000000')}</a:ln>`;
  return `<p:spPr>${serializeShapeTransform(transform)}<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${fill}${line}</p:spPr>`;
}

export function serializeShape(
  element: PptxSceneShapeElement,
  transform: PptxSceneTransform,
  shapeId: number,
): string {
  return `<p:sp>${serializeShapeNonVisualProperties(element, shapeId, false)}${serializeNativeShapeProperties(element, transform)}</p:sp>`;
}
