import type {
  PptxSceneParagraph,
  PptxSceneParagraphProperties,
} from '../scene-types';
import {
  serializeEndParagraphProperties,
  serializeTextNode,
  type PptxTextSerializationContext,
} from './text-node';

const ALIGNMENTS: Record<
  NonNullable<PptxSceneParagraphProperties['alignment']>,
  string
> = {
  center: 'ctr',
  distributed: 'dist',
  justify: 'just',
  left: 'l',
  right: 'r',
};

export function serializeParagraphProperties(
  properties: PptxSceneParagraphProperties | undefined,
): string {
  if (properties === undefined) return '';
  const attributes: string[] = [];
  if (properties.level !== undefined) {
    attributes.push(`lvl="${properties.level}"`);
  }
  if (properties.alignment !== undefined) {
    attributes.push(`algn="${ALIGNMENTS[properties.alignment]}"`);
  }
  const attributeText = attributes.length > 0 ? ` ${attributes.join(' ')}` : '';
  return `<a:pPr${attributeText}/>`;
}

export function serializeParagraph(
  paragraph: PptxSceneParagraph,
  context: PptxTextSerializationContext,
): string {
  const properties = serializeParagraphProperties(paragraph.properties);
  const children = paragraph.children
    .map((node) => serializeTextNode(node, context))
    .join('');
  const endProperties =
    paragraph.endProperties === undefined
      ? ''
      : serializeEndParagraphProperties(paragraph.endProperties);
  return `<a:p>${properties}${children}${endProperties}</a:p>`;
}
