import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';

import {
  decodeXmlEntities,
  escapeHtml,
  getTextByPathList,
  sanitizeHyperlink,
} from '../../../common';
import {
  getFontBold,
  getFontColor,
  getFontDecoration,
  getFontDecorationLine,
  getFontItalic,
  getFontShadow,
  getFontSize,
  getFontSpace,
  getFontSubscript,
  getFontType,
} from './font-style';
import {
  getHorizontalAlign,
  getParagraphIndent,
  getParagraphSpacing,
} from './paragraph';
import { serializeCssFontFamily } from '../../../common/text/css';

type ListType = 'ol' | 'ul';

interface SpanStyleInfo {
  hasLink: boolean;
  linkURL: string | null;
  styleText: string;
  text: string;
}

interface ParagraphContentNode {
  kind: 'break' | 'text';
  node: XmlLookupValue;
}

function nodeAt(
  node: unknown,
  path: readonly string[],
): XmlLookupValue | undefined {
  return getTextByPathList<XmlLookupValue>(node, path);
}

function textAt(node: unknown, path: readonly string[]): string | undefined {
  return getTextByPathList<string>(node, path);
}

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function runOrder(node: XmlLookupValue): number {
  const order = Number(
    textAt(node, ['attrs', 'order']) ?? Number.MAX_SAFE_INTEGER,
  );
  return Number.isFinite(order) ? order : Number.MAX_SAFE_INTEGER;
}

function textToHtml(value: string): string {
  return escapeHtml(decodeXmlEntities(value))
    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')
    .replace(/\s/g, '&nbsp;');
}

function renderSpan(style: SpanStyleInfo, text: string): string {
  const processedText = textToHtml(text);
  const styleAttribute = escapeHtml(style.styleText);
  if (style.hasLink && style.linkURL) {
    const href = escapeHtml(style.linkURL);
    return `<span style="${styleAttribute}"><a href="${href}" target="_blank" rel="noopener noreferrer">${processedText}</a></span>`;
  }
  return `<span style="${styleAttribute}">${processedText}</span>`;
}

export function getTextNodeValue(node: unknown): string | undefined {
  if (typeof node === 'string') return node;
  return textAt(node, ['value']);
}

function getParagraphContent(
  paragraph: XmlLookupValue,
): ParagraphContentNode[] {
  const runs = asArray(nodeAt(paragraph, ['a:r']));
  const fields = asArray(nodeAt(paragraph, ['a:fld']));
  const breaks = asArray(nodeAt(paragraph, ['a:br']));
  return [
    ...runs.map((node): ParagraphContentNode => ({ kind: 'text', node })),
    ...fields.map((node): ParagraphContentNode => ({ kind: 'text', node })),
    ...breaks.map((node): ParagraphContentNode => ({ kind: 'break', node })),
  ].sort((left, right) => runOrder(left.node) - runOrder(right.node));
}

function appendParagraphSpacing(
  styleText: string,
  spacing: ReturnType<typeof getParagraphSpacing>,
): string {
  if (!spacing) return styleText;
  if (spacing.lineSpacing !== undefined) {
    styleText += `line-height: ${spacing.lineSpacing};`;
  }
  if (spacing.spaceBefore) {
    styleText += `margin-top: ${spacing.spaceBefore};`;
  }
  if (spacing.spaceAfter) {
    styleText += `margin-bottom: ${spacing.spaceAfter};`;
  }
  return styleText;
}

function appendParagraphIndent(
  styleText: string,
  indent: ReturnType<typeof getParagraphIndent>,
  listType: ListType | '',
): string {
  if (!indent || listType) return styleText;
  if (indent.marginLeft) styleText += `margin-left: ${indent.marginLeft};`;
  if (indent.textIndent) styleText += `text-indent: ${indent.textIndent};`;
  return styleText;
}

function closeOpenLists(openLists: ListType[]): string {
  let html = '';
  while (openLists.length > 0) {
    const listType = openLists.pop();
    if (listType) html += `</li></${listType}>`;
  }
  return html;
}

export function genTextBody(
  textBodyNode: XmlLookupValue,
  shapeNode: XmlLookupValue,
  slideLayoutShape: XmlLookupValue | undefined,
  slideMasterShape: XmlLookupValue | undefined,
  type: string,
  warpObj: PptxParserContext,
): string {
  const paragraphs = asArray(nodeAt(textBodyNode, ['a:p']));
  if (paragraphs.length === 0) return '';

  const fontReference = nodeAt(shapeNode, ['p:style', 'a:fontRef']);
  const isTableCell = Boolean(nodeAt(shapeNode, ['a:tcPr']));
  const masterTextStyles = isTableCell
    ? undefined
    : warpObj.slideMasterTextStyles;
  const defaultTextStyle = isTableCell ? warpObj.defaultTextStyle : undefined;

  let html = '';
  const openLists: ListType[] = [];
  for (const paragraph of paragraphs) {
    const align = getHorizontalAlign(
      paragraph,
      shapeNode,
      type,
      slideLayoutShape,
      slideMasterShape,
      warpObj,
    );
    const spacing = getParagraphSpacing(
      paragraph,
      textBodyNode,
      slideLayoutShape,
      slideMasterShape,
      type,
      masterTextStyles,
      warpObj,
    );
    const indent = getParagraphIndent(
      paragraph,
      textBodyNode,
      slideLayoutShape,
      slideMasterShape,
      type,
      masterTextStyles,
      warpObj,
    );
    const listType = getListType(paragraph);
    const listLevel = getListLevel(paragraph);
    let styleText = `text-align: ${align};`;
    styleText = appendParagraphSpacing(styleText, spacing);
    styleText = appendParagraphIndent(styleText, indent, listType);

    if (listType) {
      const targetLevel = Math.min(listLevel, openLists.length);
      while (openLists.length > targetLevel + 1) {
        const closedList = openLists.pop();
        if (closedList) html += `</li></${closedList}>`;
      }

      if (openLists.length === targetLevel) {
        html += `<${listType}>`;
        openLists.push(listType);
      } else {
        html += '</li>';
        const currentList = openLists[targetLevel];
        if (currentList !== listType) {
          if (currentList) html += `</${currentList}>`;
          html += `<${listType}>`;
          openLists[targetLevel] = listType;
        }
      }
      html += `<li><p style="${escapeHtml(styleText)}">`;
    } else {
      html += closeOpenLists(openLists);
      html += `<p style="${escapeHtml(styleText)}">`;
    }

    const contentNodes = getParagraphContent(paragraph);
    if (contentNodes.length === 0) {
      html += genSpanElement(
        paragraph,
        paragraph,
        textBodyNode,
        fontReference,
        slideLayoutShape,
        slideMasterShape,
        type,
        masterTextStyles,
        defaultTextStyle,
        warpObj,
      );
    } else {
      let previousStyle: SpanStyleInfo | null = null;
      let accumulatedText = '';
      for (const contentNode of contentNodes) {
        if (contentNode.kind === 'break') {
          if (accumulatedText && previousStyle) {
            html += renderSpan(previousStyle, accumulatedText);
          }
          previousStyle = null;
          accumulatedText = '';
          html += '<br>';
          continue;
        }
        const style = getSpanStyleInfo(
          contentNode.node,
          paragraph,
          textBodyNode,
          fontReference,
          slideLayoutShape,
          slideMasterShape,
          type,
          masterTextStyles,
          defaultTextStyle,
          warpObj,
        );
        const startsNewSpan =
          previousStyle === null ||
          previousStyle.styleText !== style.styleText ||
          previousStyle.hasLink !== style.hasLink ||
          style.hasLink;
        if (startsNewSpan) {
          if (accumulatedText && previousStyle) {
            html += renderSpan(previousStyle, accumulatedText);
          }
          accumulatedText = '';
          if (style.hasLink) {
            html += renderSpan(style, style.text);
            previousStyle = null;
          } else {
            previousStyle = style;
            accumulatedText = style.text;
          }
        } else {
          accumulatedText += style.text;
        }
      }
      if (accumulatedText && previousStyle) {
        html += renderSpan(previousStyle, accumulatedText);
      }
    }
    html += '</p>';
  }

  html += closeOpenLists(openLists);
  return html;
}

export function getListType(node: XmlLookupValue): ListType | '' {
  const paragraphProperties = nodeAt(node, ['a:pPr']);
  if (nodeAt(paragraphProperties, ['a:buChar'])) return 'ul';
  if (nodeAt(paragraphProperties, ['a:buAutoNum'])) return 'ol';
  return '';
}

export function getListLevel(node: XmlLookupValue): number {
  const level = textAt(node, ['a:pPr', 'attrs', 'lvl']);
  switch (level) {
    case '1':
      return 1;
    case '2':
      return 2;
    case '3':
      return 3;
    case '4':
      return 4;
    case '5':
      return 5;
    case '6':
      return 6;
    case '7':
      return 7;
    case '8':
      return 8;
    default:
      return 0;
  }
}

export function genSpanElement(
  node: XmlLookupValue,
  paragraph: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  fontReference: XmlLookupValue | undefined,
  slideLayoutShape: XmlLookupValue | undefined,
  slideMasterShape: XmlLookupValue | undefined,
  type: string,
  masterTextStyles: XmlLookupValue | undefined,
  defaultTextStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): string {
  const style = getSpanStyleInfo(
    node,
    paragraph,
    textBodyNode,
    fontReference,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    defaultTextStyle,
    warpObj,
  );
  return renderSpan(style, style.text);
}

export function getSpanStyleInfo(
  node: XmlLookupValue,
  paragraph: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  fontReference: XmlLookupValue | undefined,
  slideLayoutShape: XmlLookupValue | undefined,
  slideMasterShape: XmlLookupValue | undefined,
  type: string,
  masterTextStyles: XmlLookupValue | undefined,
  defaultTextStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): SpanStyleInfo {
  const level = getListLevel(paragraph) + 1;
  const runText =
    getTextNodeValue(nodeAt(node, ['a:t'])) ??
    getTextNodeValue(nodeAt(node, ['a:fld', 'a:t'])) ??
    '\u00a0';

  const fontColor = getFontColor(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
    fontReference,
    warpObj,
  );
  const fontSize = getFontSize(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
    defaultTextStyle,
  );
  const fontType = getFontType(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
    warpObj,
  );
  const fontBold = getFontBold(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );
  const fontItalic = getFontItalic(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );
  const fontDecoration = getFontDecoration(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );
  const fontDecorationLine = getFontDecorationLine(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );
  const fontSpace = getFontSpace(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );
  const shadow = getFontShadow(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
    warpObj,
  );
  const subscript = getFontSubscript(
    node,
    paragraph,
    textBodyNode,
    slideLayoutShape,
    slideMasterShape,
    type,
    masterTextStyles,
    level,
  );

  let styleText = '';
  if (typeof fontColor === 'string') {
    if (fontColor) styleText += `color: ${fontColor};`;
  } else {
    const stops = fontColor.colors
      .map((item) => `${item.color} ${item.pos}`)
      .join(', ');
    const gradient = `linear-gradient(${fontColor.rot + 90}deg, ${stops})`;
    styleText += `background: ${gradient}; background-clip: text; color: transparent;`;
  }
  if (fontSize) styleText += `font-size: ${fontSize};`;
  const fontFamily = fontType ? serializeCssFontFamily(fontType) : null;
  if (fontFamily) styleText += `font-family: ${fontFamily};`;
  if (fontBold) styleText += `font-weight: ${fontBold};`;
  if (fontItalic) styleText += `font-style: ${fontItalic};`;
  if (fontDecoration) styleText += `text-decoration: ${fontDecoration};`;
  if (fontDecorationLine) {
    styleText += `text-decoration-line: ${fontDecorationLine};`;
  }
  if (fontSpace) styleText += `letter-spacing: ${fontSpace};`;
  if (subscript) styleText += `vertical-align: ${subscript};`;
  if (shadow) styleText += `text-shadow: ${shadow};`;

  const relationshipId = textAt(node, [
    'a:rPr',
    'a:hlinkClick',
    'attrs',
    'r:id',
  ]);
  const relationship = relationshipId
    ? warpObj.slideResObj[relationshipId]
    : undefined;
  const linkURL =
    relationship?.type === 'hyperlink'
      ? sanitizeHyperlink(relationship.target)
      : null;
  return {
    styleText,
    text: runText,
    hasLink: linkURL !== null,
    linkURL,
  };
}
