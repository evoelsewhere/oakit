import type { XmlLookupValue } from '../../../common';
import type { Fill, GradientFill, ImageFill, PatternFill } from '../types';
import type { PptxMediaData, PptxParserContext } from './context';

import tinycolor from 'tinycolor2';

import {
  angleToDegrees,
  base64ArrayBuffer,
  getMimeType,
  getTextByPathList,
  toHex,
} from '../../../common';
import {
  applyHueMod,
  applyLumMod,
  applyLumOff,
  applySatMod,
  applyShade,
  applyTint,
  getColorName2Hex,
  hslToRgb,
} from './color';
import { getSchemeColorFromTheme } from './scheme-color';

type MediaCacheKey = 'loadedAudios' | 'loadedImages' | 'loadedVideos';
type MediaMode = 'base64' | 'blob';
type GradientValue = GradientFill['value'];
type ImageValue = ImageFill['value'];
type PatternValue = PatternFill['value'];

interface PictureFilters {
  brightness?: number;
  colorTemperature?: number;
  contrast?: number;
  saturation?: number;
  sharpen?: number;
  soften?: number;
}

interface FillCandidate {
  node: XmlLookupValue;
  source: string;
}

type FillResolution =
  { state: 'found'; fill: Fill } | { state: 'missing' | 'none' };

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

function wrapChild(key: string, child: XmlLookupValue): XmlLookupValue {
  return { [key]: child } as unknown as XmlLookupValue;
}

export function getFillType(node: unknown): string {
  if (nodeAt(node, ['a:noFill'])) return 'NO_FILL';
  if (nodeAt(node, ['a:solidFill'])) return 'SOLID_FILL';
  if (nodeAt(node, ['a:gradFill'])) return 'GRADIENT_FILL';
  if (nodeAt(node, ['a:pattFill'])) return 'PATTERN_FILL';
  if (nodeAt(node, ['a:blipFill'])) return 'PIC_FILL';
  if (nodeAt(node, ['a:grpFill'])) return 'GROUP_FILL';
  return '';
}

function createImageData(ref = ''): PptxMediaData {
  return { ref, base64: '', blob: '' };
}

function createMediaData(ref = ''): Pick<PptxMediaData, 'blob' | 'ref'> {
  return { ref, blob: '' };
}

async function loadMedia(
  filePath: string,
  warpObj: PptxParserContext,
  cacheKey: MediaCacheKey,
  mode: MediaMode,
): Promise<string> {
  if (!filePath) return '';

  const normalizedPath = filePath;
  const cache = warpObj[cacheKey];
  const cacheItem = cache[normalizedPath] ?? createImageData(normalizedPath);
  cache[normalizedPath] = cacheItem;

  if (cacheItem[mode]) return cacheItem[mode];

  const fileExtension = normalizedPath.split('.').pop()?.toLowerCase() ?? '';
  if (fileExtension === 'xml') return '';

  const bytes = await warpObj.xmlReader.readMedia(normalizedPath);
  if (!bytes) return '';

  const mimeType = getMimeType(fileExtension);
  if (mode === 'base64') {
    cacheItem.base64 = `data:${mimeType};base64,${base64ArrayBuffer(bytes)}`;
  } else {
    const arrayBuffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    cacheItem.blob = URL.createObjectURL(
      new Blob([arrayBuffer], mimeType ? { type: mimeType } : undefined),
    );
  }
  return cacheItem[mode];
}

export function loadImage(
  imagePath: string,
  warpObj: PptxParserContext,
  mode: MediaMode = 'base64',
): Promise<string> {
  return loadMedia(imagePath, warpObj, 'loadedImages', mode);
}

export function loadVideo(
  videoPath: string,
  warpObj: PptxParserContext,
  mode: MediaMode = 'blob',
): Promise<string> {
  return mode === 'blob'
    ? loadMedia(videoPath, warpObj, 'loadedVideos', mode)
    : Promise.resolve('');
}

export function loadAudio(
  audioPath: string,
  warpObj: PptxParserContext,
  mode: MediaMode = 'blob',
): Promise<string> {
  return mode === 'blob'
    ? loadMedia(audioPath, warpObj, 'loadedAudios', mode)
    : Promise.resolve('');
}

export async function getImageData(
  imagePath: string,
  warpObj: PptxParserContext,
): Promise<PptxMediaData> {
  const imageData = createImageData(imagePath);
  if (!imagePath) return imageData;

  const mode = warpObj.options.imageMode;
  if (mode === 'base64' || mode === 'both') {
    imageData.base64 = await loadImage(imagePath, warpObj, 'base64');
  }
  if (mode === 'blob' || mode === 'both') {
    imageData.blob = await loadImage(imagePath, warpObj, 'blob');
  }
  return imageData;
}

export async function getVideoData(
  videoPath: string,
  warpObj: PptxParserContext,
): Promise<Pick<PptxMediaData, 'blob' | 'ref'>> {
  const videoData = createMediaData(videoPath);
  if (videoPath && warpObj.options.videoMode === 'blob') {
    videoData.blob = await loadVideo(videoPath, warpObj);
  }
  return videoData;
}

export async function getAudioData(
  audioPath: string,
  warpObj: PptxParserContext,
): Promise<Pick<PptxMediaData, 'blob' | 'ref'>> {
  const audioData = createMediaData(audioPath);
  if (audioPath && warpObj.options.audioMode === 'blob') {
    audioData.blob = await loadAudio(audioPath, warpObj);
  }
  return audioData;
}

export async function getPicFill(
  source: string,
  node: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): Promise<PptxMediaData> {
  if (!node) return createImageData();

  const relationshipId = textAt(node, ['a:blip', 'attrs', 'r:embed']);
  if (!relationshipId) return createImageData();

  const relationships =
    source === 'slideBg' || source === 'slide'
      ? warpObj.slideResObj
      : source === 'slideLayoutBg'
        ? warpObj.layoutResObj
        : source === 'slideMasterBg'
          ? warpObj.masterResObj
          : source === 'themeBg'
            ? warpObj.themeResObj
            : source === 'diagramBg'
              ? warpObj.diagramResObj
              : undefined;
  const imagePath = relationships?.[relationshipId]?.target;
  return imagePath ? getImageData(imagePath, warpObj) : createImageData();
}

export function getPicFillOpacity(node: XmlLookupValue): number {
  const amount = textAt(node, ['a:blip', 'a:alphaModFix', 'attrs', 'amt']);
  return amount ? Number.parseInt(amount) / 100_000 : 1;
}

export function getPicFilters(
  node: XmlLookupValue | undefined,
): PictureFilters | null {
  const extensions = asArray(nodeAt(node, ['a:blip', 'a:extLst', 'a:ext']));
  const filters: PictureFilters = {};

  for (const extension of extensions) {
    const effects = asArray(
      nodeAt(extension, ['a14:imgProps', 'a14:imgLayer', 'a14:imgEffect']),
    );
    for (const effect of effects) {
      const saturation = textAt(effect, ['a14:saturation', 'attrs', 'sat']);
      if (saturation)
        filters.saturation = Number.parseInt(saturation) / 100_000;

      const brightness = textAt(effect, [
        'a14:brightnessContrast',
        'attrs',
        'bright',
      ]);
      const contrast = textAt(effect, [
        'a14:brightnessContrast',
        'attrs',
        'contrast',
      ]);
      if (brightness)
        filters.brightness = Number.parseInt(brightness) / 100_000;
      if (contrast) filters.contrast = Number.parseInt(contrast) / 100_000;

      const sharpenSoften = textAt(effect, [
        'a14:sharpenSoften',
        'attrs',
        'amount',
      ]);
      if (sharpenSoften) {
        const amount = Number.parseInt(sharpenSoften) / 100_000;
        if (amount > 0) filters.sharpen = amount;
        else filters.soften = Math.abs(amount);
      }

      const colorTemperature = textAt(effect, [
        'a14:colorTemperature',
        'attrs',
        'colorTemp',
      ]);
      if (colorTemperature) {
        filters.colorTemperature = Number.parseInt(colorTemperature);
      }
    }
  }
  return Object.keys(filters).length > 0 ? filters : null;
}

async function getBackgroundPicture(
  backgroundProperties: XmlLookupValue,
  source: string,
  warpObj: PptxParserContext,
): Promise<ImageValue> {
  const picture = await getPicFill(
    source,
    nodeAt(backgroundProperties, ['a:blipFill']),
    warpObj,
  );
  const opacity = getPicFillOpacity(
    nodeAt(backgroundProperties, ['a:blipFill']) ?? backgroundProperties,
  );
  return { ...picture, opacity };
}

function normalizeGradientPath(
  path: string | undefined,
): GradientValue['path'] {
  if (path === 'circle' || path === 'rect' || path === 'shape') return path;
  return 'line';
}

function buildGradient(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  colorMap?: XmlLookupValue,
  placeholderColor?: string,
): GradientValue {
  const colors = asArray(nodeAt(node, ['a:gsLst', 'a:gs']))
    .map((stop) => {
      const position = Number(textAt(stop, ['attrs', 'pos']) ?? 0);
      return {
        pos: position ? `${position / 1000}%` : '',
        color: getSolidFill(stop, colorMap, placeholderColor, warpObj),
      };
    })
    .sort(
      (left, right) => Number.parseInt(left.pos) - Number.parseInt(right.pos),
    );

  const linearNode = nodeAt(node, ['a:lin']);
  const path = textAt(node, ['a:path', 'attrs', 'path']);
  return {
    rot: linearNode ? angleToDegrees(attributes(linearNode).ang) : 0,
    path: normalizeGradientPath(path),
    colors,
  };
}

export function getGradientFill(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): GradientValue {
  return buildGradient(node, warpObj);
}

export function getPatternFill(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): PatternValue | null {
  const pattern = nodeAt(node, ['a:pattFill']);
  if (!pattern) return null;

  return {
    type: textAt(pattern, ['attrs', 'prst']) ?? '',
    foregroundColor: nodeAt(pattern, ['a:fgClr'])
      ? getSolidFill(
          nodeAt(pattern, ['a:fgClr']),
          undefined,
          undefined,
          warpObj,
        )
      : '#000000',
    backgroundColor: nodeAt(pattern, ['a:bgClr'])
      ? getSolidFill(
          nodeAt(pattern, ['a:bgClr']),
          undefined,
          undefined,
          warpObj,
        )
      : '#FFFFFF',
  };
}

export function getBgGradientFill(
  backgroundProperties: XmlLookupValue,
  placeholderColor: string | undefined,
  slideMasterContent: XmlLookupValue,
  warpObj: PptxParserContext,
): GradientValue | string | null {
  const gradient = nodeAt(backgroundProperties, ['a:gradFill']);
  if (gradient) {
    const colorMap = nodeAt(slideMasterContent, [
      'p:sldMaster',
      'p:clrMap',
      'attrs',
    ]);
    return buildGradient(gradient, warpObj, colorMap, placeholderColor);
  }
  return placeholderColor
    ? placeholderColor.startsWith('#')
      ? placeholderColor
      : `#${placeholderColor}`
    : null;
}

function getMasterColorMap(
  context: PptxParserContext,
): XmlLookupValue | undefined {
  return nodeAt(context.slideMasterContent, [
    'p:sldMaster',
    'p:clrMap',
    'attrs',
  ]);
}

function getOverrideColorMap(
  content: XmlLookupValue,
  rootName: 'p:sld' | 'p:sldLayout',
): XmlLookupValue | undefined {
  return nodeAt(content, [
    rootName,
    'p:clrMapOvr',
    'a:overrideClrMapping',
    'attrs',
  ]);
}

function orderedThemeBackgroundFills(
  warpObj: PptxParserContext,
): XmlLookupValue[] {
  const fillList = nodeAt(warpObj.themeContent, [
    'a:theme',
    'a:themeElements',
    'a:fmtScheme',
    'a:bgFillStyleLst',
  ]);
  if (!fillList) return [];

  const fills: { node: XmlLookupValue; order: number }[] = [];
  for (const key of Object.keys(fillList)) {
    if (key === 'attrs') continue;
    for (const child of asArray(nodeAt(fillList, [key]))) {
      fills.push({
        node: wrapChild(key, child),
        order: Number(attributes(child).order ?? Number.MAX_SAFE_INTEGER),
      });
    }
  }
  return fills
    .sort((left, right) => left.order - right.order)
    .map(({ node }) => node);
}

async function resolveBackgroundNode(
  node: XmlLookupValue,
  source: string,
  colorMap: XmlLookupValue | undefined,
  placeholderColor: string | undefined,
  warpObj: PptxParserContext,
): Promise<Fill | null> {
  const fillType = getFillType(node);
  if (fillType === 'SOLID_FILL') {
    return {
      type: 'color',
      value: getSolidFill(
        nodeAt(node, ['a:solidFill']),
        colorMap,
        placeholderColor,
        warpObj,
      ),
    };
  }
  if (fillType === 'GRADIENT_FILL') {
    const gradient = nodeAt(node, ['a:gradFill']);
    return gradient
      ? {
          type: 'gradient',
          value: buildGradient(gradient, warpObj, colorMap, placeholderColor),
        }
      : null;
  }
  if (fillType === 'PIC_FILL') {
    return {
      type: 'image',
      value: await getBackgroundPicture(node, source, warpObj),
    };
  }
  if (fillType === 'PATTERN_FILL') {
    const pattern = getPatternFill(node, warpObj);
    return pattern ? { type: 'pattern', value: pattern } : null;
  }
  return null;
}

interface BackgroundLevel {
  colorMap: XmlLookupValue | undefined;
  content: XmlLookupValue;
  rootName: 'p:sld' | 'p:sldLayout' | 'p:sldMaster';
  source: string;
}

export async function getSlideBackgroundFill(
  warpObj: PptxParserContext,
): Promise<Fill> {
  const masterColorMap = getMasterColorMap(warpObj);
  const layoutColorMap =
    getOverrideColorMap(warpObj.slideLayoutContent, 'p:sldLayout') ??
    masterColorMap;
  const slideColorMap =
    getOverrideColorMap(warpObj.slideContent, 'p:sld') ?? layoutColorMap;
  const levels: BackgroundLevel[] = [
    {
      content: warpObj.slideContent,
      rootName: 'p:sld',
      source: 'slideBg',
      colorMap: slideColorMap,
    },
    {
      content: warpObj.slideLayoutContent,
      rootName: 'p:sldLayout',
      source: 'slideLayoutBg',
      colorMap: layoutColorMap,
    },
    {
      content: warpObj.slideMasterContent,
      rootName: 'p:sldMaster',
      source: 'slideMasterBg',
      colorMap: masterColorMap,
    },
  ];

  for (const level of levels) {
    const backgroundProperties = nodeAt(level.content, [
      level.rootName,
      'p:cSld',
      'p:bg',
      'p:bgPr',
    ]);
    if (backgroundProperties) {
      return (
        (await resolveBackgroundNode(
          backgroundProperties,
          level.source,
          level.colorMap,
          undefined,
          warpObj,
        )) ?? { type: 'color', value: '#fff' }
      );
    }

    const backgroundReference = nodeAt(level.content, [
      level.rootName,
      'p:cSld',
      'p:bg',
      'p:bgRef',
    ]);
    if (backgroundReference) {
      const placeholderColor = getSolidFill(
        backgroundReference,
        level.colorMap,
        undefined,
        warpObj,
      );
      const index = Number(attributes(backgroundReference).idx ?? 0) - 1000;
      const themeFill =
        index > 0 ? orderedThemeBackgroundFills(warpObj)[index - 1] : undefined;
      return themeFill
        ? ((await resolveBackgroundNode(
            themeFill,
            'themeBg',
            level.colorMap,
            placeholderColor,
            warpObj,
          )) ?? { type: 'color', value: '#fff' })
        : { type: 'color', value: '#fff' };
    }
  }
  return { type: 'color', value: '#fff' };
}

function getShapeFillCandidates(
  node: XmlLookupValue,
  source: string,
  slideLayoutShape: XmlLookupValue | undefined,
  slideMasterShape: XmlLookupValue | undefined,
): FillCandidate[] {
  const candidates: FillCandidate[] = [{ node, source }];
  if (slideLayoutShape) {
    candidates.push({ node: slideLayoutShape, source: 'slideLayoutBg' });
  }
  if (slideMasterShape) {
    candidates.push({ node: slideMasterShape, source: 'slideMasterBg' });
  }
  return candidates;
}

async function findFillInGroupHierarchy(
  groupHierarchy: XmlLookupValue[],
  warpObj: PptxParserContext,
  source: string,
): Promise<Fill | null> {
  for (const groupNode of groupHierarchy) {
    const groupProperties = nodeAt(groupNode, ['p:grpSpPr']);
    if (!groupProperties) continue;

    const result = await resolveShapeProperties(
      groupProperties,
      warpObj,
      source,
      [],
    );
    if (result.state === 'found') return result.fill;
    if (result.state === 'none') return null;
  }
  return null;
}

async function resolveShapeProperties(
  shapeProperties: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  groupHierarchy: XmlLookupValue[],
): Promise<FillResolution> {
  const fillType = getFillType(shapeProperties);
  if (fillType === 'NO_FILL') return { state: 'none' };
  if (fillType === 'SOLID_FILL') {
    const value = getSolidFill(
      nodeAt(shapeProperties, ['a:solidFill']),
      undefined,
      undefined,
      warpObj,
    );
    return value
      ? { state: 'found', fill: { type: 'color', value } }
      : { state: 'missing' };
  }
  if (fillType === 'GRADIENT_FILL') {
    const gradient = nodeAt(shapeProperties, ['a:gradFill']);
    return gradient
      ? {
          state: 'found',
          fill: { type: 'gradient', value: getGradientFill(gradient, warpObj) },
        }
      : { state: 'missing' };
  }
  if (fillType === 'PIC_FILL') {
    const pictureNode = nodeAt(shapeProperties, ['a:blipFill']);
    const picture = await getPicFill(source, pictureNode, warpObj);
    return pictureNode
      ? {
          state: 'found',
          fill: {
            type: 'image',
            value: { ...picture, opacity: getPicFillOpacity(pictureNode) },
          },
        }
      : { state: 'missing' };
  }
  if (fillType === 'PATTERN_FILL') {
    const pattern = getPatternFill(shapeProperties, warpObj);
    return pattern
      ? { state: 'found', fill: { type: 'pattern', value: pattern } }
      : { state: 'missing' };
  }
  if (fillType === 'GROUP_FILL') {
    const fill = await findFillInGroupHierarchy(
      groupHierarchy,
      warpObj,
      source,
    );
    return fill ? { state: 'found', fill } : { state: 'none' };
  }
  return { state: 'missing' };
}

async function resolveShapeFillFromNode(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  groupHierarchy: XmlLookupValue[],
): Promise<FillResolution> {
  const shapeProperties = nodeAt(node, ['p:spPr']);
  if (shapeProperties) {
    const result = await resolveShapeProperties(
      shapeProperties,
      warpObj,
      source,
      groupHierarchy,
    );
    if (result.state !== 'missing') return result;
  }

  const fillReference = nodeAt(node, ['p:style', 'a:fillRef']);
  const value = getSolidFill(fillReference, undefined, undefined, warpObj);
  return value
    ? { state: 'found', fill: { type: 'color', value } }
    : { state: 'missing' };
}

export async function getShapeFill(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  source: string,
  options: {
    groupHierarchy?: XmlLookupValue[];
    slideLayoutSpNode?: XmlLookupValue | undefined;
    slideMasterSpNode?: XmlLookupValue | undefined;
  } = {},
): Promise<Fill | null> {
  const { groupHierarchy = [], slideLayoutSpNode, slideMasterSpNode } = options;
  const candidates = getShapeFillCandidates(
    node,
    source,
    slideLayoutSpNode,
    slideMasterSpNode,
  );
  for (const candidate of candidates) {
    const result = await resolveShapeFillFromNode(
      candidate.node,
      warpObj,
      candidate.source,
      groupHierarchy,
    );
    if (result.state === 'none') return null;
    if (result.state === 'found') return result.fill;
  }
  return null;
}

function modifier(node: XmlLookupValue | undefined, name: string): number {
  const value = textAt(node, [name, 'attrs', 'val']);
  return value === undefined ? Number.NaN : Number.parseInt(value) / 100_000;
}

function percentComponent(value: string | undefined): number {
  return Number(value?.replace('%', '') ?? 0) / 100;
}

export function getSolidFill(
  solidFill: XmlLookupValue | undefined,
  colorMap: XmlLookupValue | undefined,
  placeholderColor: string | undefined,
  warpObj: PptxParserContext,
): string {
  if (!solidFill) return '';

  let color = '';
  let colorNode: XmlLookupValue | undefined;
  if ((colorNode = nodeAt(solidFill, ['a:srgbClr']))) {
    color = textAt(colorNode, ['attrs', 'val']) ?? '';
  } else if ((colorNode = nodeAt(solidFill, ['a:schemeClr']))) {
    const scheme = textAt(colorNode, ['attrs', 'val']) ?? '';
    color =
      getSchemeColorFromTheme(
        `a:${scheme}`,
        warpObj,
        colorMap,
        placeholderColor,
      ) ?? '';
  } else if ((colorNode = nodeAt(solidFill, ['a:scrgbClr']))) {
    const values = attributes(colorNode);
    color =
      toHex(255 * percentComponent(values.r)) +
      toHex(255 * percentComponent(values.g)) +
      toHex(255 * percentComponent(values.b));
  } else if ((colorNode = nodeAt(solidFill, ['a:prstClr']))) {
    color = getColorName2Hex(textAt(colorNode, ['attrs', 'val']) ?? '') ?? '';
  } else if ((colorNode = nodeAt(solidFill, ['a:hslClr']))) {
    const values = attributes(colorNode);
    const rgb = hslToRgb(
      Number(values.hue ?? 0) / 100_000,
      percentComponent(values.sat),
      percentComponent(values.lum),
    );
    color = toHex(rgb.r) + toHex(rgb.g) + toHex(rgb.b);
  } else if ((colorNode = nodeAt(solidFill, ['a:sysClr']))) {
    color = textAt(colorNode, ['attrs', 'lastClr']) ?? '';
  }

  let hasAlpha = false;
  const alpha = modifier(colorNode, 'a:alpha');
  if (!Number.isNaN(alpha)) {
    color = tinycolor(color).setAlpha(alpha).toHex8();
    hasAlpha = true;
  }

  const transformations: [
    string,
    (value: string, amount: number, alpha: boolean) => string,
  ][] = [
    ['a:hueMod', applyHueMod],
    ['a:lumMod', applyLumMod],
    ['a:lumOff', applyLumOff],
    ['a:satMod', applySatMod],
    ['a:shade', applyShade],
    ['a:tint', applyTint],
  ];
  for (const [name, transform] of transformations) {
    const amount = modifier(colorNode, name);
    if (!Number.isNaN(amount)) color = transform(color, amount, hasAlpha);
  }
  return color && !color.startsWith('#') ? `#${color}` : color;
}
