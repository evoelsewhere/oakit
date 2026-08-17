import { decodeXmlEntities } from '../../../common/text/html';
import { escapeXmlText } from '../writer/xml';
import { unsupportedPptxEdit } from './patch-error';
import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditableShapeXml,
} from './shape-range';
import type { PptxRoundTripReplaceTextOperation } from './types';

const OFFICE_ESCAPE_PATTERN = /_x[0-9a-f]{4}_/i;

export function patchPptxShapeTextXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripReplaceTextOperation,
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
      'PowerPoint text edit target contains unsupported compatibility markup',
    );
  }
  if (
    pptxShapeHasElement(shape, qualifiedPptxName(drawingPrefix, 'br')) ||
    pptxShapeHasElement(shape, qualifiedPptxName(drawingPrefix, 'fld'))
  ) {
    unsupportedPptxEdit(
      'PowerPoint text edit target must contain one plain text run',
    );
  }

  const textName = qualifiedPptxName(drawingPrefix, 't');
  const escapedTextName = escapePptxXmlPattern(textName);
  const textPattern = new RegExp(
    `<${escapedTextName}((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*>([^<]*)<\\/${escapedTextName}\\s*>`,
    'g',
  );
  const matches = [...shape.matchAll(textPattern)];
  if (matches.length !== 1) {
    unsupportedPptxEdit(
      'PowerPoint text edit target must contain exactly one text node',
    );
  }
  const match = matches[0] as RegExpMatchArray;
  const sourceText = decodeXmlEntities(match[2] as string);
  if (
    OFFICE_ESCAPE_PATTERN.test(sourceText) ||
    sourceText !== operation.expectedText
  ) {
    unsupportedPptxEdit(
      'PowerPoint text edit source XML does not match its preview precondition',
    );
  }
  const originalAttributes = match[1] as string;
  const attributesWithoutSpace = originalAttributes.replace(
    /\s+xml:space\s*=\s*(?:"[^"]*"|'[^']*')/gi,
    '',
  );
  const replacement = `<${textName}${attributesWithoutSpace} xml:space="preserve">${escapeXmlText(operation.value)}</${textName}>`;
  const matchStart = range.start + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacement}${xml.slice(matchEnd)}`;
}
