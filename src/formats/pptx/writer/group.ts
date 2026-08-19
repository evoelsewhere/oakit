import type {
  PptxSceneGroupElement,
  PptxSceneGroupTransform,
} from '../scene-types';
import { degreesToAngle, pointsToEmu } from './units';
import { escapeXmlAttribute } from './xml';

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function serializeGroupNonVisualProperties(
  element: PptxSceneGroupElement,
  shapeId: number,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `Group ${shapeId}`)}"`,
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
  return `<p:nvGrpSpPr><p:cNvPr ${attributes.join(' ')}/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`;
}

export function serializeGroupTransform(
  transform: PptxSceneGroupTransform,
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
  const attributeText =
    attributes.length === 0 ? '' : ` ${attributes.join(' ')}`;
  const child = transform.childSpace;
  return `<a:xfrm${attributeText}><a:off x="${pointsToEmu(transform.x)}" y="${pointsToEmu(transform.y)}"/><a:ext cx="${pointsToEmu(transform.width)}" cy="${pointsToEmu(transform.height)}"/><a:chOff x="${pointsToEmu(child.x)}" y="${pointsToEmu(child.y)}"/><a:chExt cx="${pointsToEmu(child.width)}" cy="${pointsToEmu(child.height)}"/></a:xfrm>`;
}

export function serializeGroup(
  element: PptxSceneGroupElement,
  transform: PptxSceneGroupTransform,
  shapeId: number,
  childrenXml: string,
): string {
  return `<p:grpSp>${serializeGroupNonVisualProperties(element, shapeId)}<p:grpSpPr>${serializeGroupTransform(transform)}</p:grpSpPr>${childrenXml}</p:grpSp>`;
}
