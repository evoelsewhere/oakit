import type { PptxSceneElement, PptxSceneSlide } from '../scene-types';
import { serializeSolidColorFill } from './color';
import type { PptxTextSerializationContext } from './text-node';
import { serializeShape } from './shape';
import { serializeTextShape } from './text-shape';
import { escapeXmlAttribute } from './xml';
import { serializePicture } from './image';
import { serializeTable } from './table';
import { serializeGroup } from './group';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const SHAPE_TREE_ROOT =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function serializeElement(
  element: PptxSceneElement,
  context: PptxTextSerializationContext,
  imageRelationships: ReadonlyMap<string, string>,
  allocateShapeId: () => number,
): string {
  if (element.type === 'unsupported') {
    throw new TypeError('PowerPoint slide writer rejects opaque elements');
  }
  const transform = element.authored.transform;
  if (transform === undefined) {
    throw new TypeError(
      `PowerPoint slide writer requires an authored ${element.type} transform`,
    );
  }
  const shapeId = allocateShapeId();
  switch (element.type) {
    case 'group': {
      if (!('childSpace' in transform)) {
        throw new TypeError(
          `PowerPoint group element ${element.key} requires a child-space transform`,
        );
      }
      const children = element.elements
        .map((child) =>
          serializeElement(child, context, imageRelationships, allocateShapeId),
        )
        .join('');
      return serializeGroup(element, transform, shapeId, children);
    }
    case 'text':
      return serializeTextShape(element, transform, shapeId, context);
    case 'shape':
      return serializeShape(element, transform, shapeId);
    case 'image': {
      const relationshipId = imageRelationships.get(element.key);
      if (relationshipId === undefined) {
        throw new TypeError(
          `PowerPoint image element ${element.key} has no media relationship`,
        );
      }
      return serializePicture(element, transform, shapeId, relationshipId);
    }
    case 'table':
      return serializeTable(element, transform, shapeId, context);
  }
}

export function serializeSlide(
  slide: PptxSceneSlide,
  context: PptxTextSerializationContext,
  imageRelationships: ReadonlyMap<string, string> = new Map(),
): string {
  const rootAttributes = [
    `xmlns:a="${DRAWING_NAMESPACE}"`,
    `xmlns:r="${RELATIONSHIPS_NAMESPACE}"`,
    `xmlns:p="${PRESENTATION_NAMESPACE}"`,
  ];
  if (slide.hidden !== undefined) {
    rootAttributes.push(`show="${slide.hidden ? '0' : '1'}"`);
  }
  const commonSlideName =
    slide.name === undefined ? '' : ` name="${escapeXmlAttribute(slide.name)}"`;
  const background =
    slide.backgroundColor === undefined
      ? ''
      : `<p:bg><p:bgPr>${serializeSolidColorFill(slide.backgroundColor)}<a:effectLst/></p:bgPr></p:bg>`;
  let nextShapeId = 2;
  const allocateShapeId = (): number => {
    const result = nextShapeId;
    nextShapeId += 1;
    return result;
  };
  const elements = slide.elements
    .map((element) =>
      serializeElement(element, context, imageRelationships, allocateShapeId),
    )
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${rootAttributes.join(' ')}><p:cSld${commonSlideName}>${background}<p:spTree>${SHAPE_TREE_ROOT}${elements}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
