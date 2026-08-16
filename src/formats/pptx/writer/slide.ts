import type { PptxSceneElement, PptxSceneSlide } from '../scene-types';
import type { PptxTextSerializationContext } from './text-node';
import { serializeTextShape } from './text-shape';
import { escapeXmlAttribute } from './xml';

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
  shapeId: number,
  context: PptxTextSerializationContext,
): string {
  if (element.type !== 'text') {
    throw new TypeError('PowerPoint slide writer accepts text elements only');
  }
  const transform = element.authored.transform;
  if (transform === undefined) {
    throw new TypeError(
      'PowerPoint slide writer requires an authored text transform',
    );
  }
  return serializeTextShape(element, transform, shapeId, context);
}

export function serializeSlide(
  slide: PptxSceneSlide,
  context: PptxTextSerializationContext,
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
  const elements = slide.elements
    .map((element, index) => serializeElement(element, index + 2, context))
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${rootAttributes.join(' ')}><p:cSld${commonSlideName}><p:spTree>${SHAPE_TREE_ROOT}${elements}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
