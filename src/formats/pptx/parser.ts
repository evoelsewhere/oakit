import type { XmlLookupValue } from '../../common';
import type {
  PptxMediaData,
  PptxNodeIndex,
  PptxParserContext,
  PptxRelationshipMap,
} from './internal/context';
import type {
  Element,
  Audio,
  Diagram,
  Group,
  Image,
  Math as MathElement,
  PptxDocument,
  PptxDiagnostic,
  PptxInput,
  PptxParseOptions,
  Shape,
  Slide,
  Table,
  TableCell,
  Text,
  Video,
} from './types';

type Draft<T extends { id: string }> = Omit<T, 'id'> & { id?: string };
type ElementDraft = Element extends infer T
  ? T extends { id: string }
    ? Draft<T>
    : never
  : never;

import JSZip from 'jszip';
import { getBorder } from './internal/border';
import {
  getSlideBackgroundFill,
  getShapeFill,
  getSolidFill,
  getPicFill,
  getPicFilters,
  getImageData,
  getVideoData,
  getAudioData,
} from './internal/fill';
import { getChartInfo } from './internal/chart';
import { getVerticalAlign, getTextAutoFit } from './internal/paragraph';
import { getTextInsets } from './internal/text-insets';
import { getPosition, getSize } from './internal/position';
import { genTextBody, getTextNodeValue } from './internal/text';
import {
  getCustomShapePath,
  identifyShape,
  isStrokeOnlyCustomGeometry,
} from './internal/shape';
import {
  extractFileExtension,
  getTextByPathList,
  angleToDegrees,
  isVideoLink,
  escapeHtml,
  getRelationshipPartUri,
  hasValidText,
  numberToFixed,
  resolveRelationshipTarget,
} from '../../common';
import { getShadow } from './internal/shadow';
import {
  getTableBorders,
  getTableCellParams,
  getTableRowParams,
} from './internal/table';
import { RATIO_EMUs_Points } from '../../common/ooxml/units';
import { findOMath, latexFormart, parseOMath } from './internal/math';
import { getShapePath } from './internal/shape-path';
import { parseTransition, findTransitionNode } from './internal/animation';
import { getDiagramNodeContext, getSmartArtTextData } from './internal/diagram';
import { PptxXmlReader } from './internal/xml-reader';

function nodeAt(
  node: unknown,
  path: readonly string[],
): XmlLookupValue | undefined {
  return getTextByPathList<XmlLookupValue>(node, path);
}

function textAt(node: unknown, path: readonly string[]): string | undefined {
  return getTextByPathList<string>(node, path);
}

function attributes(node: unknown): Record<string, string> {
  return getTextByPathList<Record<string, string>>(node, ['attrs']) ?? {};
}

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function emptyXmlNode(): XmlLookupValue {
  return {} as unknown as XmlLookupValue;
}

export async function parse(
  file: PptxInput,
  options: PptxParseOptions = {},
  diagnostics: PptxDiagnostic[] = [],
): Promise<PptxDocument> {
  const slides: Slide[] = [];
  const loadedImages: Record<string, PptxMediaData> = {};
  const loadedVideos: Record<string, PptxMediaData> = {};
  const loadedAudios: Record<string, PptxMediaData> = {};
  const parseOptions: Required<PptxParseOptions> = {
    ...options,
    imageMode: options.imageMode || 'base64',
    videoMode: options.videoMode || 'none',
    audioMode: options.audioMode || 'none',
    errorMode: options.errorMode || 'tolerant',
  };

  const zip = await JSZip.loadAsync(file);
  const xmlReader = new PptxXmlReader(zip, parseOptions.errorMode, diagnostics);

  const filesInfo = await getContentTypes(xmlReader);
  const { width, height, defaultTextStyle } = await getSlideInfo(xmlReader);
  const { themeContent, themeColors } = await getTheme(xmlReader);
  const usedFonts = await getUsedFonts(xmlReader);

  for (const filename of filesInfo.slides) {
    const singleSlide = await processSingleSlide(
      zip,
      xmlReader,
      filename,
      themeContent,
      defaultTextStyle,
      loadedImages,
      loadedVideos,
      loadedAudios,
      parseOptions,
    );
    slides.push(singleSlide);
  }

  return {
    slides,
    usedFonts,
    themeColors,
    size: {
      width,
      height,
    },
  };
}

async function getContentTypes(xmlReader: PptxXmlReader) {
  const contentTypes = await xmlReader.read('[Content_Types].xml', {
    required: true,
  });
  const overrides = asArray(nodeAt(contentTypes, ['Types', 'Override']));
  const slides: string[] = [];
  const slideLayouts: string[] = [];

  for (const item of overrides) {
    const itemAttributes = attributes(item);
    const partName = itemAttributes.PartName?.replace(/^\//, '');
    if (!partName) continue;
    switch (itemAttributes.ContentType) {
      case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
        slides.push(partName);
        break;
      case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
        slideLayouts.push(partName);
        break;
      default:
    }
  }

  const slideNumber = (filename: string): number => {
    const match = /(\d+)\.xml$/.exec(filename);
    return Number(match?.[1] ?? Number.MAX_SAFE_INTEGER);
  };
  return {
    slides: slides.sort(
      (left, right) => slideNumber(left) - slideNumber(right),
    ),
    slideLayouts: slideLayouts.sort(
      (left, right) => slideNumber(left) - slideNumber(right),
    ),
  };
}

async function getUsedFonts(xmlReader: PptxXmlReader): Promise<string[]> {
  const content = await xmlReader.read('ppt/presentation.xml', {
    required: true,
  });
  const embeddedFonts = asArray(
    nodeAt(content, ['p:presentation', 'p:embeddedFontLst', 'p:embeddedFont']),
  );
  const usedFonts: string[] = [];
  for (const embeddedFont of embeddedFonts) {
    const typeface = textAt(embeddedFont, ['p:font', 'attrs', 'typeface']);
    if (typeface && !usedFonts.includes(typeface)) usedFonts.push(typeface);
  }

  return usedFonts;
}

async function getSlideInfo(xmlReader: PptxXmlReader) {
  const content = await xmlReader.read('ppt/presentation.xml', {
    required: true,
  });
  const sizeAttributes = attributes(
    nodeAt(content, ['p:presentation', 'p:sldSz']),
  );
  const defaultTextStyle =
    nodeAt(content, ['p:presentation', 'p:defaultTextStyle']) ?? emptyXmlNode();
  return {
    width: Number(sizeAttributes.cx ?? 0) * RATIO_EMUs_Points,
    height: Number(sizeAttributes.cy ?? 0) * RATIO_EMUs_Points,
    defaultTextStyle,
  };
}

async function getTheme(xmlReader: PptxXmlReader) {
  const presentationPart = 'ppt/presentation.xml';
  const themeRelationship = (
    await getRelationships(xmlReader, presentationPart)
  ).find(
    (relationship) =>
      attributes(relationship).Type ===
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  );
  let themeContent: XmlLookupValue | null = null;
  const themeAttributes = attributes(themeRelationship);
  const themeUri = themeAttributes.Target;
  if (themeUri) {
    const themeFilename = resolveRelationshipTarget(
      presentationPart,
      themeUri,
      themeAttributes.TargetMode,
    );
    themeContent = await xmlReader.read(themeFilename);
  }

  const themeColors: string[] = [];
  const colorScheme = nodeAt(themeContent, [
    'a:theme',
    'a:themeElements',
    'a:clrScheme',
  ]);
  if (colorScheme) {
    for (let i = 1; i <= 6; i++) {
      const accent = nodeAt(colorScheme, [`a:accent${i}`]);
      if (!accent) break;
      const color = textAt(accent, ['a:srgbClr', 'attrs', 'val']);
      if (color) themeColors.push('#' + color);
    }
  }

  return { themeContent, themeColors };
}

const STANDARD_RELATIONSHIP_PREFIX =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

async function getRelationships(
  xmlReader: PptxXmlReader,
  ownerPart: string,
): Promise<XmlLookupValue[]> {
  if (!ownerPart) return [];
  const content = await xmlReader.read(getRelationshipPartUri(ownerPart));
  return asArray(nodeAt(content, ['Relationships', 'Relationship']));
}

function addRelationship(
  relationships: PptxRelationshipMap,
  relationship: XmlLookupValue,
  target: string,
): void {
  const values = attributes(relationship);
  if (!values.Id || !values.Type) return;
  relationships[values.Id] = {
    type: values.Type.replace(STANDARD_RELATIONSHIP_PREFIX, ''),
    target,
  };
}

async function processSingleSlide(
  zip: JSZip,
  xmlReader: PptxXmlReader,
  slideFilename: string,
  themeContent: XmlLookupValue | null,
  defaultTextStyle: XmlLookupValue,
  loadedImages: Record<string, PptxMediaData>,
  loadedVideos: Record<string, PptxMediaData>,
  loadedAudios: Record<string, PptxMediaData>,
  options: Required<PptxParseOptions>,
): Promise<Slide> {
  const slideRelationships = await getRelationships(xmlReader, slideFilename);

  let noteFilename = '';
  let layoutFilename = '';
  let masterFilename = '';
  let themeFilename = '';
  const slideResObj: PptxRelationshipMap = {};
  const layoutResObj: PptxRelationshipMap = {};
  const masterResObj: PptxRelationshipMap = {};
  const themeResObj: PptxRelationshipMap = {};

  for (const relationship of slideRelationships) {
    const values = attributes(relationship);
    if (!values.Target || !values.Type) continue;
    const target = resolveRelationshipTarget(
      slideFilename,
      values.Target,
      values.TargetMode,
    );
    if (values.Type === `${STANDARD_RELATIONSHIP_PREFIX}slideLayout`) {
      layoutFilename = target;
    } else if (values.Type === `${STANDARD_RELATIONSHIP_PREFIX}notesSlide`) {
      noteFilename = target;
    }
    addRelationship(slideResObj, relationship, target);
  }

  const slideNotesContent = noteFilename
    ? await xmlReader.read(noteFilename)
    : emptyXmlNode();
  const note = getNote(slideNotesContent);

  const slideLayoutContent = await xmlReader.read(layoutFilename);
  const slideLayoutTables = indexNodes(slideLayoutContent);
  const slideLayoutResFilename = layoutFilename
    ? getRelationshipPartUri(layoutFilename)
    : '';
  const slideLayoutResContent = await xmlReader.read(slideLayoutResFilename);
  const layoutRelationships = asArray(
    nodeAt(slideLayoutResContent, ['Relationships', 'Relationship']),
  );
  for (const relationship of layoutRelationships) {
    const values = attributes(relationship);
    if (!values.Target || !values.Type) continue;
    const target = resolveRelationshipTarget(
      layoutFilename,
      values.Target,
      values.TargetMode,
    );
    if (values.Type === `${STANDARD_RELATIONSHIP_PREFIX}slideMaster`) {
      masterFilename = target;
    } else {
      addRelationship(layoutResObj, relationship, target);
    }
  }

  const slideMasterContent = await xmlReader.read(masterFilename);
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, [
    'p:sldMaster',
    'p:txStyles',
  ]);
  const slideMasterTables = indexNodes(slideMasterContent);
  const slideMasterResFilename = masterFilename
    ? getRelationshipPartUri(masterFilename)
    : '';
  const slideMasterResContent = await xmlReader.read(slideMasterResFilename);
  const masterRelationships = asArray(
    nodeAt(slideMasterResContent, ['Relationships', 'Relationship']),
  );
  for (const relationship of masterRelationships) {
    const values = attributes(relationship);
    if (!values.Target || !values.Type) continue;
    const target = resolveRelationshipTarget(
      masterFilename,
      values.Target,
      values.TargetMode,
    );
    if (values.Type === `${STANDARD_RELATIONSHIP_PREFIX}theme`) {
      themeFilename = target;
    } else {
      addRelationship(masterResObj, relationship, target);
    }
  }

  let currentThemeContent = themeContent;
  if (!currentThemeContent && themeFilename) {
    currentThemeContent = await xmlReader.read(themeFilename);
  }
  if (themeFilename) {
    const themeName = themeFilename.split('/').pop();
    if (themeName) {
      for (const relationship of await getRelationships(
        xmlReader,
        themeFilename,
      )) {
        const values = attributes(relationship);
        if (!values.Target) continue;
        addRelationship(
          themeResObj,
          relationship,
          resolveRelationshipTarget(
            themeFilename,
            values.Target,
            values.TargetMode,
          ),
        );
      }
    }
  }

  const tableStyles = await xmlReader.read('ppt/tableStyles.xml');

  const slideContent = await xmlReader.read(slideFilename, { required: true });
  const nodes = nodeAt(slideContent, ['p:sld', 'p:cSld', 'p:spTree']);
  const warpObj: PptxParserContext = {
    zip,
    loadedImages,
    loadedVideos,
    loadedAudios,
    options,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    tableStyles,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent: currentThemeContent ?? emptyXmlNode(),
    themeResObj,
    diagramFileCache: {},
    defaultTextStyle,
    xmlReader,
  };
  const layoutElements = await getLayoutElements(warpObj);
  const fill = await getSlideBackgroundFill(warpObj);

  const elements: Element[] = [];
  for (const nodeKey of Object.keys(nodes ?? {})) {
    for (const node of asArray(nodeAt(nodes, [nodeKey]))) {
      const ret = await processNodesInSlide(nodeKey, node, warpObj, 'slide');
      if (ret) elements.push(ret);
    }
  }

  let transitionNode = findTransitionNode(slideContent, 'p:sld');
  if (!transitionNode)
    transitionNode = findTransitionNode(slideLayoutContent, 'p:sldLayout');
  if (!transitionNode)
    transitionNode = findTransitionNode(slideMasterContent, 'p:sldMaster');

  const transition = parseTransition(transitionNode);

  return {
    fill,
    elements,
    layoutElements,
    note,
    transition,
  };
}

function getHyperlinkFromCNvPr(
  cNvPr: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  const linkId = textAt(cNvPr, ['a:hlinkClick', 'attrs', 'r:id']);
  if (!linkId) return null;

  const res = warpObj['slideResObj'][linkId];
  if (!res) return null;

  if (res['type'] !== 'hyperlink') return null;

  const target = res['target'];
  if (!target || !/^https?:\/\//.test(target)) return null;

  return target;
}

function getNote(noteContent: XmlLookupValue): string {
  let text = '';
  const shapeNodes = asArray(
    nodeAt(noteContent, ['p:notes', 'p:cSld', 'p:spTree', 'p:sp']),
  );
  for (const shapeNode of shapeNodes) {
    const placeholderType = textAt(shapeNode, [
      'p:nvSpPr',
      'p:nvPr',
      'p:ph',
      'attrs',
      'type',
    ]);
    if (placeholderType !== 'body') continue;

    const textBody = nodeAt(shapeNode, ['p:txBody']);
    if (!textBody) continue;

    const listTypes: (string | undefined)[] = [];
    for (const paragraph of asArray(nodeAt(textBody, ['a:p']))) {
      const paragraphProperties = nodeAt(paragraph, ['a:pPr']);
      const alignment = textAt(paragraphProperties, ['attrs', 'algn']);
      let align = 'left';
      if (alignment) {
        switch (alignment) {
          case 'r':
            align = 'right';
            break;
          case 'ctr':
            align = 'center';
            break;
          case 'just':
          case 'dist':
            align = 'justify';
            break;
          default:
            break;
        }
      }

      let listType = '';
      if (nodeAt(paragraphProperties, ['a:buChar'])) listType = 'ul';
      else if (nodeAt(paragraphProperties, ['a:buAutoNum'])) listType = 'ol';
      const level = textAt(paragraphProperties, ['attrs', 'lvl']);
      const listLevel = level === undefined ? 0 : Number.parseInt(level);

      if (listType) {
        while (listTypes.length > listLevel + 1) {
          text += `</${listTypes.pop()}>`;
        }
        if (listTypes[listLevel] === undefined) {
          text += `<${listType}>`;
          listTypes[listLevel] = listType;
        } else if (listTypes[listLevel] !== listType) {
          text += `</${listTypes[listLevel]}>`;
          text += `<${listType}>`;
          listTypes[listLevel] = listType;
        }
        text += `<li><p style="text-align:${align};">`;
      } else {
        while (listTypes.length > 0) {
          text += `</${listTypes.pop()}>`;
        }
        text += `<p style="text-align:${align};">`;
      }

      for (const run of asArray(nodeAt(paragraph, ['a:r']))) {
        const value = getTextNodeValue(nodeAt(run, ['a:t']));
        if (value) text += value;
      }

      if (listType) text += '</p></li>';
      else text += '</p>';
    }
    while (listTypes.length > 0) {
      text += `</${listTypes.pop()}>`;
    }
  }
  return text;
}

async function getLayoutElements(
  warpObj: PptxParserContext,
): Promise<Element[]> {
  const elements: Element[] = [];
  const layoutTree = nodeAt(warpObj.slideLayoutContent, [
    'p:sldLayout',
    'p:cSld',
    'p:spTree',
  ]);
  const masterTree = nodeAt(warpObj.slideMasterContent, [
    'p:sldMaster',
    'p:cSld',
    'p:spTree',
  ]);
  const showMasterShapes = textAt(warpObj.slideLayoutContent, [
    'p:sldLayout',
    'attrs',
    'showMasterSp',
  ]);

  async function appendTree(
    tree: XmlLookupValue | undefined,
    source: string,
  ): Promise<void> {
    for (const nodeKey of Object.keys(tree ?? {})) {
      for (const node of asArray(nodeAt(tree, [nodeKey]))) {
        if (nodeAt(node, ['p:nvSpPr', 'p:nvPr', 'p:ph'])) continue;
        const element = await processNodesInSlide(
          nodeKey,
          node,
          warpObj,
          source,
        );
        if (element) elements.push(element);
      }
    }
  }

  await appendTree(layoutTree, 'slideLayoutBg');
  if (showMasterShapes !== '0') await appendTree(masterTree, 'slideMasterBg');
  return elements;
}

function indexNodes(content: XmlLookupValue): PptxNodeIndex {
  const rootKey = Object.keys(content)[0];
  const shapeTree = rootKey
    ? nodeAt(content, [rootKey, 'p:cSld', 'p:spTree'])
    : undefined;
  const idTable: Record<string, XmlLookupValue> = {};
  const idxTable: Record<string, XmlLookupValue> = {};
  const typeTable: Record<string, XmlLookupValue> = {};

  for (const key of Object.keys(shapeTree ?? {})) {
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue;
    for (const targetNode of asArray(nodeAt(shapeTree, [key]))) {
      const nonVisualProperties = nodeAt(targetNode, ['p:nvSpPr']);
      const id = textAt(nonVisualProperties, ['p:cNvPr', 'attrs', 'id']);
      const index = textAt(nonVisualProperties, [
        'p:nvPr',
        'p:ph',
        'attrs',
        'idx',
      ]);
      const type = textAt(nonVisualProperties, [
        'p:nvPr',
        'p:ph',
        'attrs',
        'type',
      ]);
      if (id) idTable[id] = targetNode;
      if (index) idxTable[index] = targetNode;
      if (type && !typeTable[type]) typeTable[type] = targetNode;
    }
  }

  return { idTable, idxTable, typeTable };
}

async function processNodesInSlide(
  nodeKey: string,
  nodeValue: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  groupHierarchy: XmlLookupValue[] = [],
): Promise<Element | null> {
  let json: ElementDraft | null = null;

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = await processSpNode(nodeValue, warpObj, source, groupHierarchy);
      break;
    case 'p:cxnSp': // Shape, Text
      json = await processCxnSpNode(nodeValue, warpObj, source);
      break;
    case 'p:pic': // Image, Video, Audio
      json = await processPicNode(nodeValue, warpObj, source);
      break;
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source);
      break;
    case 'p:grpSp':
      json = await processGroupSpNode(
        nodeValue,
        warpObj,
        source,
        groupHierarchy,
      );
      break;
    case 'mc:AlternateContent':
      if (nodeAt(nodeValue, ['mc:Fallback', 'p:grpSpPr', 'a:xfrm'])) {
        const fallback = nodeAt(nodeValue, ['mc:Fallback']);
        if (!fallback) break;
        json = await processGroupSpNode(
          fallback,
          warpObj,
          source,
          groupHierarchy,
        );
      } else if (nodeAt(nodeValue, ['mc:Choice'])) {
        json = await processMathNode(nodeValue, warpObj, source);
      }
      break;
    default:
  }

  if (json && !json.id) {
    const id =
      getTextByPathList(nodeValue, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id']) ||
      getTextByPathList(nodeValue, ['p:nvPicPr', 'p:cNvPr', 'attrs', 'id']) ||
      getTextByPathList(nodeValue, ['p:nvCxnSpPr', 'p:cNvPr', 'attrs', 'id']) ||
      getTextByPathList(nodeValue, ['p:nvGrpSpPr', 'p:cNvPr', 'attrs', 'id']) ||
      getTextByPathList(nodeValue, [
        'p:nvGraphicFramePr',
        'p:cNvPr',
        'attrs',
        'id',
      ]) ||
      getTextByPathList(nodeValue, [
        'mc:Choice',
        'p:sp',
        'p:nvSpPr',
        'p:cNvPr',
        'attrs',
        'id',
      ]) ||
      getTextByPathList(nodeValue, [
        'mc:Fallback',
        'p:sp',
        'p:nvSpPr',
        'p:cNvPr',
        'attrs',
        'id',
      ]) ||
      getTextByPathList(nodeValue, [
        'mc:Fallback',
        'p:nvGrpSpPr',
        'p:cNvPr',
        'attrs',
        'id',
      ]);

    json.id = id || '';
  }
  return json as Element | null;
}

async function processMathNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
): Promise<Draft<MathElement> | null> {
  const choice = nodeAt(node, ['mc:Choice']);
  const fallback = nodeAt(node, ['mc:Fallback']);
  const order = Number(attributes(node).order ?? 0);
  const xfrmNode = nodeAt(choice, ['p:sp', 'p:spPr', 'a:xfrm']);
  const { top, left } = getPosition(xfrmNode, undefined, undefined);
  const { width, height } = getSize(xfrmNode, undefined, undefined);

  const oMath = findOMath(choice)[0];
  if (!oMath) return null;

  const latex = latexFormart(parseOMath(oMath));

  const blipFill = nodeAt(fallback, ['p:sp', 'p:spPr', 'a:blipFill']);
  const picFill = await getPicFill(source, blipFill, warpObj);

  let text = '';
  if (nodeAt(choice, ['p:sp', 'p:txBody', 'a:p', 'a:r'])) {
    const sp = nodeAt(choice, ['p:sp']);
    const textBody = nodeAt(sp, ['p:txBody']);
    if (sp && textBody) {
      text = genTextBody(textBody, sp, undefined, undefined, '', warpObj);
    }
  }

  return {
    type: 'math',
    top,
    left,
    width,
    height,
    latex,
    picRef: picFill.ref,
    picBase64: picFill.base64,
    picBlob: picFill.blob,
    text,
    order,
  };
}

async function processGroupSpNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  parentGroupHierarchy: XmlLookupValue[] = [],
): Promise<Draft<Group> | null> {
  const order = Number(attributes(node).order ?? 0);
  const xfrmNode = nodeAt(node, ['p:grpSpPr', 'a:xfrm']);
  if (!xfrmNode) return null;
  const x =
    Number(textAt(xfrmNode, ['a:off', 'attrs', 'x']) ?? 0) * RATIO_EMUs_Points;
  const y =
    Number(textAt(xfrmNode, ['a:off', 'attrs', 'y']) ?? 0) * RATIO_EMUs_Points;
  const childX =
    Number(textAt(xfrmNode, ['a:chOff', 'attrs', 'x']) ?? 0) *
    RATIO_EMUs_Points;
  const childY =
    Number(textAt(xfrmNode, ['a:chOff', 'attrs', 'y']) ?? 0) *
    RATIO_EMUs_Points;
  const width =
    Number(textAt(xfrmNode, ['a:ext', 'attrs', 'cx']) ?? 0) * RATIO_EMUs_Points;
  const height =
    Number(textAt(xfrmNode, ['a:ext', 'attrs', 'cy']) ?? 0) * RATIO_EMUs_Points;
  const childWidth =
    Number(textAt(xfrmNode, ['a:chExt', 'attrs', 'cx']) ?? 0) *
    RATIO_EMUs_Points;
  const childHeight =
    Number(textAt(xfrmNode, ['a:chExt', 'attrs', 'cy']) ?? 0) *
    RATIO_EMUs_Points;
  const isFlipV = textAt(xfrmNode, ['attrs', 'flipV']) === '1';
  const isFlipH = textAt(xfrmNode, ['attrs', 'flipH']) === '1';
  const rotate = angleToDegrees(textAt(xfrmNode, ['attrs', 'rot']));

  const ws = childWidth === 0 ? 0 : width / childWidth;
  const hs = childHeight === 0 ? 0 : height / childHeight;
  const currentGroupHierarchy = [...parentGroupHierarchy, node];
  const elements: Element[] = [];
  for (const nodeKey of Object.keys(node)) {
    for (const child of asArray(nodeAt(node, [nodeKey]))) {
      const ret = await processNodesInSlide(
        nodeKey,
        child,
        warpObj,
        source,
        currentGroupHierarchy,
      );
      if (ret) elements.push(ret);
    }
  }

  const transformGroupedElement = (
    element: Element,
    offsetX = 0,
    offsetY = 0,
  ): Element => {
    const elementRotate = 'rotate' in element ? element.rotate : 0;
    const normalizedRotate = ((elementRotate % 360) + 360) % 360;
    const isUniformScale = Math.abs(ws - hs) < 0.000001;
    const shouldSwapDimensions =
      normalizedRotate === 90 || normalizedRotate === 270;
    const centerX = element.left + element.width / 2;
    const centerY = element.top + element.height / 2;
    const nextCenterX = (centerX - offsetX) * ws;
    const nextCenterY = (centerY - offsetY) * hs;
    const widthScale = shouldSwapDimensions && !isUniformScale ? hs : ws;
    const heightScale = shouldSwapDimensions && !isUniformScale ? ws : hs;
    const width = element.width * widthScale;
    const height = element.height * heightScale;

    return {
      ...element,
      left: numberToFixed(nextCenterX - width / 2),
      top: numberToFixed(nextCenterY - height / 2),
      width: numberToFixed(width),
      height: numberToFixed(height),
    };
  };

  const processedElements = elements.map((element) => {
    const transformed = transformGroupedElement(element, childX, childY);
    return transformed.type === 'group'
      ? {
          ...transformed,
          elements: processNestedGroupElements(transformed.elements),
        }
      : transformed;
  });

  function processNestedGroupElements(
    nestedElements: Element[],
    depth = 0,
  ): Element[] {
    if (depth > 10) return nestedElements;
    return nestedElements.map((element) => {
      const processed = transformGroupedElement(element);
      if (processed.type === 'group') {
        return {
          ...processed,
          elements: processNestedGroupElements(processed.elements, depth + 1),
        };
      }
      return processed;
    });
  }

  return {
    type: 'group',
    top: numberToFixed(y),
    left: numberToFixed(x),
    width: numberToFixed(width),
    height: numberToFixed(height),
    rotate,
    order,
    isFlipV,
    isFlipH,
    elements: processedElements,
  };
}

async function processSpNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  groupHierarchy: XmlLookupValue[] = [],
): Promise<Draft<Shape> | Draft<Text>> {
  const nonVisualProperties = nodeAt(node, ['p:nvSpPr', 'p:cNvPr']);
  const name = textAt(nonVisualProperties, ['attrs', 'name']) ?? '';
  const index = textAt(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
  let type = textAt(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);

  let slideLayoutSpNode: XmlLookupValue | undefined;
  let slideMasterSpNode: XmlLookupValue | undefined;

  if (type) {
    if (index) {
      slideLayoutSpNode = warpObj.slideLayoutTables.idxTable[index];
      slideMasterSpNode = warpObj.slideMasterTables.idxTable[index];
      if (!slideLayoutSpNode)
        slideLayoutSpNode = warpObj.slideLayoutTables.typeTable[type];
      if (!slideMasterSpNode)
        slideMasterSpNode = warpObj.slideMasterTables.typeTable[type];
    } else {
      slideLayoutSpNode = warpObj.slideLayoutTables.typeTable[type];
      slideMasterSpNode = warpObj.slideMasterTables.typeTable[type];
    }
  } else if (index) {
    slideLayoutSpNode = warpObj.slideLayoutTables.idxTable[index];
    slideMasterSpNode = warpObj.slideMasterTables.idxTable[index];
  }

  if (!type) {
    const textBox = textAt(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox']);
    if (textBox === '1') type = 'text';
  }
  if (!type)
    type = textAt(slideLayoutSpNode, [
      'p:nvSpPr',
      'p:nvPr',
      'p:ph',
      'attrs',
      'type',
    ]);
  if (!type)
    type = textAt(slideMasterSpNode, [
      'p:nvSpPr',
      'p:nvPr',
      'p:ph',
      'attrs',
      'type',
    ]);
  if (!slideMasterSpNode && type === 'ctrTitle')
    slideMasterSpNode = warpObj.slideMasterTables.typeTable.title;

  if (!type) {
    if (source === 'diagramBg') type = 'diagram';
    else type = 'obj';
  }

  const link = getHyperlinkFromCNvPr(nonVisualProperties, warpObj);

  return await genShape(
    node,
    slideLayoutSpNode,
    slideMasterSpNode,
    name,
    type,
    order,
    warpObj,
    source,
    link,
    groupHierarchy,
  );
}

async function processCxnSpNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
): Promise<Draft<Shape> | Draft<Text>> {
  const nonVisualProperties = nodeAt(node, ['p:nvCxnSpPr', 'p:cNvPr']);
  const name = textAt(nonVisualProperties, ['attrs', 'name']) ?? '';
  const type = textAt(node, ['p:nvCxnSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);
  const link = getHyperlinkFromCNvPr(nonVisualProperties, warpObj);

  return await genShape(
    node,
    undefined,
    undefined,
    name,
    type,
    order,
    warpObj,
    source,
    link,
  );
}

async function genShape(
  node: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  name: string,
  type: string | undefined,
  order: number,
  warpObj: PptxParserContext,
  source: string,
  link: string | null,
  groupHierarchy: XmlLookupValue[] = [],
): Promise<Draft<Shape> | Draft<Text>> {
  const xfrmList = ['p:spPr', 'a:xfrm'];
  const slideXfrmNode = nodeAt(node, xfrmList);
  const slideLayoutXfrmNode = nodeAt(slideLayoutSpNode, xfrmList);
  const slideMasterXfrmNode = nodeAt(slideMasterSpNode, xfrmList);

  const shapeType = textAt(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst']);
  const customGeometry = nodeAt(node, ['p:spPr', 'a:custGeom']);

  const keypoints: Record<string, number> = {};
  if (shapeType) {
    const adjustmentNodes = asArray(
      nodeAt(node, ['p:spPr', 'a:prstGeom', 'a:avLst', 'a:gd']),
    );
    for (const adjustment of adjustmentNodes) {
      const adjustmentName = textAt(adjustment, ['attrs', 'name']);
      const formula = textAt(adjustment, ['attrs', 'fmla']);
      if (adjustmentName && formula?.startsWith('val ')) {
        keypoints[adjustmentName] = Number(formula.slice(4)) / 50_000;
      }
    }
  }

  const { top, left } = getPosition(
    slideXfrmNode,
    slideLayoutXfrmNode,
    slideMasterXfrmNode,
  );
  const { width, height } = getSize(
    slideXfrmNode,
    slideLayoutXfrmNode,
    slideMasterXfrmNode,
  );
  const pathViewBox = { x: 0, y: 0, width, height };

  const isFlipV = textAt(slideXfrmNode, ['attrs', 'flipV']) === '1';
  const isFlipH = textAt(slideXfrmNode, ['attrs', 'flipH']) === '1';

  const rotate = angleToDegrees(textAt(slideXfrmNode, ['attrs', 'rot']));

  const textTransform = nodeAt(node, ['p:txXfrm']);
  const textRotationValue = textAt(textTransform, ['attrs', 'rot']);
  const textRotation = textRotationValue
    ? angleToDegrees(textRotationValue) + 90
    : rotate;

  let content = '';
  const textBody = nodeAt(node, ['p:txBody']);
  if (textBody)
    content = genTextBody(
      textBody,
      node,
      slideLayoutSpNode,
      slideMasterSpNode,
      type ?? '',
      warpObj,
    );

  const {
    borderColor,
    borderWidth,
    borderType,
    strokeDasharray,
    headEnd,
    tailEnd,
  } = getBorder(node, type, warpObj);
  const fill = await getShapeFill(node, warpObj, source, {
    groupHierarchy,
    slideLayoutSpNode,
    slideMasterSpNode,
  });

  const outerShadow = nodeAt(node, ['p:spPr', 'a:effectLst', 'a:outerShdw']);
  const shadow = outerShadow ? getShadow(outerShadow, warpObj) : undefined;

  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode);
  const isVertical =
    textAt(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert';
  const wrap =
    textAt(node, ['p:txBody', 'a:bodyPr', 'attrs', 'wrap']) !== 'none';
  const autoFit = getTextAutoFit(node, slideLayoutSpNode, slideMasterSpNode);
  const textInset = getTextInsets(node, slideLayoutSpNode, slideMasterSpNode);

  const commonData = {
    left,
    top,
    width,
    height,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fill,
    content,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    wrap,
    name,
    order,
    ...(shadow ? { shadow } : {}),
    ...(autoFit ? { autoFit } : {}),
    ...(link ? { link } : {}),
    ...(textInset ? { textInset } : {}),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };

  const hasText = Boolean(content && hasValidText(content));
  if (customGeometry && type !== 'diagram') {
    const extension = attributes(nodeAt(slideXfrmNode, ['a:ext']));
    const customWidth = Number(extension.cx ?? 0) * RATIO_EMUs_Points;
    const customHeight = Number(extension.cy ?? 0) * RATIO_EMUs_Points;
    const customShapeData: Draft<Shape> = {
      ...commonData,
      type: 'shape',
      shapType: 'custom',
      content: hasText ? content : '',
      path: getCustomShapePath(customGeometry, customWidth, customHeight),
      pathViewBox: {
        x: 0,
        y: 0,
        width: customWidth,
        height: customHeight,
      },
      ...(isStrokeOnlyCustomGeometry(customGeometry)
        ? { strokeOnly: true }
        : {}),
    };
    return customShapeData;
  }

  const shapePath = shapeType
    ? getShapePath(shapeType, width, height, node)
    : '';
  const STROKE_ONLY_PRESET_SHAPE_TYPES = [
    'arc',
    'leftBrace',
    'rightBrace',
    'bracePair',
    'leftBracket',
    'rightBracket',
    'bracketPair',
  ];
  const isStrokeOnlyPresetShape =
    shapeType !== undefined &&
    STROKE_ONLY_PRESET_SHAPE_TYPES.includes(shapeType);

  if (shapeType && (type === 'obj' || !type || shapeType !== 'rect')) {
    return {
      ...commonData,
      type: 'shape',
      content: hasText ? content : '',
      shapType: shapeType,
      path: shapePath,
      pathViewBox,
      keypoints,
      ...(isStrokeOnlyPresetShape ? { strokeOnly: true } : {}),
    };
  }
  if (shapeType && !hasText && (fill || borderWidth)) {
    return {
      ...commonData,
      type: 'shape',
      content: '',
      shapType: shapeType,
      path: shapePath,
      pathViewBox,
      keypoints,
      ...(isStrokeOnlyPresetShape ? { strokeOnly: true } : {}),
    };
  }
  return {
    ...commonData,
    type: 'text',
    isVertical,
    rotate: textRotation,
  };
}

async function processPicNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
): Promise<Draft<Audio> | Draft<Image> | Draft<Video> | null> {
  const relationships =
    source === 'slideMasterBg'
      ? warpObj.masterResObj
      : source === 'slideLayoutBg'
        ? warpObj.layoutResObj
        : warpObj.slideResObj;
  const nonVisualProperties = nodeAt(node, ['p:nvPicPr', 'p:cNvPr']);
  const link = getHyperlinkFromCNvPr(nonVisualProperties, warpObj);
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);
  const relationshipId = textAt(node, [
    'p:blipFill',
    'a:blip',
    'attrs',
    'r:embed',
  ]);
  const imageName = relationshipId
    ? relationships[relationshipId]?.target
    : undefined;
  if (!imageName) return null;

  let xfrmNode = nodeAt(node, ['p:spPr', 'a:xfrm']);
  if (!xfrmNode) {
    const index = textAt(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    if (index) {
      xfrmNode = nodeAt(warpObj.slideLayoutTables.idxTable[index], [
        'p:spPr',
        'a:xfrm',
      ]);
    }
  }

  const { top, left } = getPosition(xfrmNode, undefined, undefined);
  const { width, height } = getSize(xfrmNode, undefined, undefined);
  const imageData = await getImageData(imageName, warpObj);

  const isFlipV = textAt(xfrmNode, ['attrs', 'flipV']) === '1';
  const isFlipH = textAt(xfrmNode, ['attrs', 'flipH']) === '1';
  const rotate = angleToDegrees(textAt(xfrmNode, ['attrs', 'rot']));

  const videoNode = nodeAt(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile']);
  let videoData: Pick<PptxMediaData, 'blob' | 'ref'> | undefined;
  if (videoNode) {
    const videoRelationshipId = textAt(videoNode, ['attrs', 'r:link']);
    const videoFile = videoRelationshipId
      ? relationships[videoRelationshipId]?.target
      : undefined;
    if (videoFile && isVideoLink(videoFile)) {
      videoData = { ref: escapeHtml(videoFile), blob: '' };
    } else if (videoFile) {
      const extension = extractFileExtension(videoFile).toLowerCase();
      if (extension === 'mp4' || extension === 'webm' || extension === 'ogg') {
        videoData = await getVideoData(videoFile, warpObj);
      } else {
        videoData = { ref: videoFile, blob: '' };
      }
    }
  }

  const audioNode = nodeAt(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile']);
  let audioData: Pick<PptxMediaData, 'blob' | 'ref'> | undefined;
  if (audioNode) {
    const audioRelationshipId = textAt(audioNode, ['attrs', 'r:link']);
    const audioFile = audioRelationshipId
      ? relationships[audioRelationshipId]?.target
      : undefined;
    const extension = audioFile
      ? extractFileExtension(audioFile).toLowerCase()
      : '';
    if (
      audioFile &&
      (extension === 'mp3' || extension === 'wav' || extension === 'ogg')
    ) {
      audioData = await getAudioData(audioFile, warpObj);
    } else if (audioFile) {
      audioData = { ref: audioFile, blob: '' };
    }
  }

  if (videoData) {
    return {
      type: 'video',
      top,
      left,
      width,
      height,
      rotate,
      ref: videoData.ref,
      blob: videoData.blob,
      order,
    };
  }
  if (audioData) {
    return {
      type: 'audio',
      top,
      left,
      width,
      height,
      rotate,
      ref: audioData.ref,
      blob: audioData.blob,
      order,
    };
  }

  const sourceRectangle = getTextByPathList<Record<string, string>>(node, [
    'p:blipFill',
    'a:srcRect',
    'attrs',
  ]);
  const rect: Image['rect'] | undefined = sourceRectangle
    ? {
        ...(sourceRectangle.t ? { t: Number(sourceRectangle.t) / 1000 } : {}),
        ...(sourceRectangle.b ? { b: Number(sourceRectangle.b) / 1000 } : {}),
        ...(sourceRectangle.l ? { l: Number(sourceRectangle.l) / 1000 } : {}),
        ...(sourceRectangle.r ? { r: Number(sourceRectangle.r) / 1000 } : {}),
      }
    : undefined;
  let geom = 'rect';
  const presetGeometry = textAt(node, [
    'p:spPr',
    'a:prstGeom',
    'attrs',
    'prst',
  ]);
  const customGeometry = nodeAt(node, ['p:spPr', 'a:custGeom']);

  if (presetGeometry) {
    geom = presetGeometry;
  } else if (customGeometry) {
    geom = identifyShape(customGeometry);
    if (geom !== 'custom') geom = `custom:${geom}`;
  }

  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(
    node,
    undefined,
    warpObj,
  );

  const filters = getPicFilters(nodeAt(node, ['p:blipFill']));
  return {
    type: 'image',
    top,
    left,
    width,
    height,
    rotate,
    ref: imageData.ref,
    base64: imageData.base64,
    blob: imageData.blob,
    isFlipV,
    isFlipH,
    order,
    ...(rect && Object.keys(rect).length > 0 ? { rect } : {}),
    geom,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    ...(filters ? { filters } : {}),
    ...(link ? { link } : {}),
  };
}

async function processGraphicFrameNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
): Promise<ElementDraft | null> {
  const graphicTypeUri = textAt(node, [
    'a:graphic',
    'a:graphicData',
    'attrs',
    'uri',
  ]);

  let result: ElementDraft | null = null;
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = await genTable(node, warpObj);
      break;
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj);
      break;
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = await genDiagram(node, warpObj);
      break;
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole': {
      let oleObject = nodeAt(node, [
        'a:graphic',
        'a:graphicData',
        'mc:AlternateContent',
        'mc:Fallback',
        'p:oleObj',
      ]);
      if (!oleObject)
        oleObject = nodeAt(node, ['a:graphic', 'a:graphicData', 'p:oleObj']);
      if (oleObject)
        result = await processGroupSpNode(oleObject, warpObj, source);
      break;
    }
    default:
  }
  return result;
}

async function genTable(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): Promise<Draft<Table>> {
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);
  const tableNode = nodeAt(node, ['a:graphic', 'a:graphicData', 'a:tbl']);
  const xfrmNode = nodeAt(node, ['p:xfrm']);
  const { top, left } = getPosition(xfrmNode, undefined, undefined);
  const { width, height } = getSize(xfrmNode, undefined, undefined);

  const tableProperties = nodeAt(node, [
    'a:graphic',
    'a:graphicData',
    'a:tbl',
    'a:tblPr',
  ]);
  const columnNodes = asArray(
    nodeAt(node, [
      'a:graphic',
      'a:graphicData',
      'a:tbl',
      'a:tblGrid',
      'a:gridCol',
    ]),
  );
  const colWidths = columnNodes.map(
    (column) => Number(textAt(column, ['attrs', 'w']) ?? 0) * RATIO_EMUs_Points,
  );

  const tablePropertyAttributes = attributes(tableProperties);
  const tableStyleAttributes = {
    isFrstRowAttr: tablePropertyAttributes.firstRow === '1' ? 1 : 0,
    isFrstColAttr: tablePropertyAttributes.firstCol === '1' ? 1 : 0,
    isLstRowAttr: tablePropertyAttributes.lastRow === '1' ? 1 : 0,
    isLstColAttr: tablePropertyAttributes.lastCol === '1' ? 1 : 0,
    isBandRowAttr: tablePropertyAttributes.bandRow === '1' ? 1 : 0,
    isBandColAttr: tablePropertyAttributes.bandCol === '1' ? 1 : 0,
  };

  const tableStyleId = getTextNodeValue(
    nodeAt(tableProperties, ['a:tableStyleId']),
  );
  const tableStyle = tableStyleId
    ? asArray(
        nodeAt(warpObj.tableStyles, ['a:tblStyleLst', 'a:tblStyle']),
      ).find((style) => attributes(style).styleId === tableStyleId)
    : undefined;

  const borderStyle = nodeAt(tableStyle, [
    'a:wholeTbl',
    'a:tcStyle',
    'a:tcBdr',
  ]);
  const borders: Table['borders'] = borderStyle
    ? getTableBorders(borderStyle, warpObj)
    : {};

  const backgroundFill =
    nodeAt(tableStyle, ['a:tblBg', 'a:fillRef']) ??
    nodeAt(tableStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill']);
  const tableBackground = getSolidFill(
    backgroundFill,
    undefined,
    undefined,
    warpObj,
  );

  const rowNodes = asArray(nodeAt(tableNode, ['a:tr']));
  const data: TableCell[][] = [];
  const rowHeights: number[] = [];
  for (const [rowIndex, rowNode] of rowNodes.entries()) {
    rowHeights.push(
      Number(textAt(rowNode, ['attrs', 'h']) ?? 0) * RATIO_EMUs_Points,
    );

    const { fillColor, fontColor, fontBold } = getTableRowParams(
      rowNodes,
      rowIndex,
      tableStyleAttributes,
      tableStyle,
      warpObj,
    );
    const cells: TableCell[] = [];
    const cellNodes = asArray(nodeAt(rowNode, ['a:tc']));
    for (const [columnIndex, cellNode] of cellNodes.entries()) {
      let cellSource: string | undefined;
      const isFirstColumn = columnIndex === 0;
      const isLastColumn = columnIndex === cellNodes.length - 1;
      const isFirstRow = rowIndex === 0;
      const isLastRow = rowIndex === rowNodes.length - 1;
      if (isFirstColumn && tableStyleAttributes.isFrstColAttr === 1) {
        cellSource = 'a:firstCol';
        if (isLastRow && nodeAt(tableStyle, ['a:seCell']))
          cellSource = 'a:seCell';
        else if (isFirstRow && nodeAt(tableStyle, ['a:neCell']))
          cellSource = 'a:neCell';
      } else if (
        columnIndex > 0 &&
        tableStyleAttributes.isBandColAttr === 1 &&
        !isLastColumn
      ) {
        cellSource =
          columnIndex % 2 === 0 && nodeAt(tableStyle, ['a:band1V'])
            ? 'a:band1V'
            : nodeAt(tableStyle, ['a:band2V'])
              ? 'a:band2V'
              : undefined;
      }
      if (isLastColumn && tableStyleAttributes.isLstColAttr === 1) {
        cellSource = 'a:lastCol';
        if (isLastRow && nodeAt(tableStyle, ['a:swCell']))
          cellSource = 'a:swCell';
        else if (isFirstRow && nodeAt(tableStyle, ['a:nwCell']))
          cellSource = 'a:nwCell';
      }

      const textBody = nodeAt(cellNode, ['a:txBody']);
      const text = textBody
        ? genTextBody(textBody, cellNode, undefined, undefined, '', warpObj)
        : '';
      const cell = await getTableCellParams(
        cellNode,
        tableStyle,
        cellSource,
        warpObj,
      );
      const resolvedFontBold = cell.fontBold ?? fontBold;
      const resolvedFontColor = cell.fontColor || fontColor;
      const resolvedFillColor = cell.fillColor || fillColor || tableBackground;
      cells.push({
        text,
        vAlign: cell.vAlign,
        borders: cell.borders,
        ...(cell.rowSpan ? { rowSpan: cell.rowSpan } : {}),
        ...(cell.colSpan ? { colSpan: cell.colSpan } : {}),
        ...(cell.vMerge ? { vMerge: cell.vMerge } : {}),
        ...(cell.hMerge ? { hMerge: cell.hMerge } : {}),
        ...(resolvedFontBold !== undefined
          ? { fontBold: resolvedFontBold }
          : {}),
        ...(resolvedFontColor ? { fontColor: resolvedFontColor } : {}),
        ...(resolvedFillColor ? { fillColor: resolvedFillColor } : {}),
      });
    }
    data.push(cells);
  }

  const measuredWidth = colWidths.reduce((sum, value) => sum + value, 0);

  return {
    type: 'table',
    top,
    left,
    width: measuredWidth ? numberToFixed(measuredWidth) : width,
    height,
    data,
    order,
    borders,
    rowHeights,
    colWidths,
  };
}

async function genChart(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): Promise<ElementDraft | null> {
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);
  const xfrmNode = nodeAt(node, ['p:xfrm']);
  const { top, left } = getPosition(xfrmNode, undefined, undefined);
  const { width, height } = getSize(xfrmNode, undefined, undefined);

  const relationshipId = textAt(node, [
    'a:graphic',
    'a:graphicData',
    'c:chart',
    'attrs',
    'r:id',
  ]);
  const referenceName = relationshipId
    ? (warpObj.slideResObj[relationshipId]?.target ??
      warpObj.layoutResObj[relationshipId]?.target ??
      warpObj.masterResObj[relationshipId]?.target)
    : undefined;
  if (!referenceName) return null;

  const content = await warpObj.xmlReader.read(referenceName);
  const plotArea = nodeAt(content, ['c:chartSpace', 'c:chart', 'c:plotArea']);

  if (!plotArea) return null;
  const chart = getChartInfo(plotArea, warpObj);

  if (!chart) return null;

  return {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    colors: chart.colors,
    chartType: chart.type,
    order,
    ...(chart.marker !== undefined ? { marker: chart.marker } : {}),
    ...(chart.barDir !== undefined ? { barDir: chart.barDir } : {}),
    ...(chart.holeSize !== undefined ? { holeSize: chart.holeSize } : {}),
    ...(chart.grouping !== undefined ? { grouping: chart.grouping } : {}),
    ...(chart.style !== undefined ? { style: chart.style } : {}),
  } as ElementDraft;
}

async function genDiagram(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): Promise<Draft<Diagram>> {
  const order = Number(textAt(node, ['attrs', 'order']) ?? 0);
  const xfrmNode = nodeAt(node, ['p:xfrm']);
  const { left, top } = getPosition(xfrmNode, undefined, undefined);
  const { width, height } = getSize(xfrmNode, undefined, undefined);

  const diagramWarpObj = await getDiagramNodeContext(node, warpObj);
  const diagramShapes = asArray(
    nodeAt(diagramWarpObj.digramFileContent, ['p:drawing', 'p:spTree', 'p:sp']),
  );
  const elements: (Shape | Text)[] = [];
  for (const shapeNode of diagramShapes) {
    const element = await processNodesInSlide(
      'p:sp',
      shapeNode,
      diagramWarpObj,
      'diagramBg',
    );
    if (element?.type === 'shape' || element?.type === 'text') {
      elements.push(element);
    }
  }
  const diagramData = diagramWarpObj.diagramContent?.data;
  const textList = diagramData ? getSmartArtTextData(diagramData) : [];

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
    textList,
    order,
  };
}
