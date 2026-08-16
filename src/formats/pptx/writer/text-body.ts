import type {
  PptxSceneTextBody,
  PptxSceneTextBodyProperties,
} from '../scene-types';
import { serializeParagraph } from './paragraph';
import type { PptxTextSerializationContext } from './text-node';

const ANCHORS: Record<
  NonNullable<PptxSceneTextBodyProperties['anchor']>,
  string
> = {
  bottom: 'b',
  center: 'ctr',
  distributed: 'dist',
  justified: 'just',
  top: 't',
};

const AUTO_FIT_ELEMENTS: Record<
  NonNullable<PptxSceneTextBodyProperties['autoFit']>,
  string
> = {
  none: '<a:noAutofit/>',
  shape: '<a:spAutoFit/>',
  text: '<a:normAutofit/>',
};

export function serializeTextBodyProperties(
  properties: PptxSceneTextBodyProperties,
): string {
  const attributes: string[] = [];
  if (properties.anchor !== undefined) {
    attributes.push(`anchor="${ANCHORS[properties.anchor]}"`);
  }
  if (properties.vertical !== undefined) {
    attributes.push(`vert="${properties.vertical ? 'eaVert' : 'horz'}"`);
  }
  if (properties.wrap !== undefined) {
    attributes.push(`wrap="${properties.wrap ? 'square' : 'none'}"`);
  }
  const attributeText = attributes.length > 0 ? ` ${attributes.join(' ')}` : '';
  if (properties.autoFit === undefined) {
    return `<a:bodyPr${attributeText}/>`;
  }
  return `<a:bodyPr${attributeText}>${AUTO_FIT_ELEMENTS[properties.autoFit]}</a:bodyPr>`;
}

export function serializeTextBody(
  text: PptxSceneTextBody,
  context: PptxTextSerializationContext,
): string {
  const bodyProperties = serializeTextBodyProperties(text.body);
  const paragraphs = text.paragraphs
    .map((paragraph) => serializeParagraph(paragraph, context))
    .join('');
  return `<p:txBody>${bodyProperties}<a:lstStyle/>${paragraphs}</p:txBody>`;
}
