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
import { normalizeHexColor } from '../../../common/text/css';

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

interface AlphaTransformation {
  apply: (current: number, amount: number) => number;
  kind: 'alpha';
  name: string;
}

interface ColorValueTransformation {
  apply: (value: string, amount: number, hasAlpha: boolean) => string;
  kind: 'color';
  name: string;
}

type ColorTransformation = AlphaTransformation | ColorValueTransformation;
type AuthoredColorTransformation = ColorTransformation & {
  amount: number;
  order: number;
};

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

function finiteNumber(value: string | undefined): number | undefined {
  if (value === undefined || value.trim() === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function fixedPercentage(value: string | undefined): number | undefined {
  const parsed = finiteNumber(value?.replace(/%$/, ''));
  if (parsed === undefined) return undefined;
  return parsed / (value?.endsWith('%') ? 100 : 100_000);
}

function clampUnit(value: number): number {
  return Math.min(Math.max(value, 0), 1);
}

function byteHex(value: number): string {
  const finite = Number.isFinite(value) ? value : 0;
  return toHex(Math.round(Math.min(Math.max(finite, 0), 255)));
}

function normalizedColor(value: string | undefined): string {
  if (value === undefined) return '';
  return normalizeHexColor(value) ?? '';
}

function replaceAlpha(_current: number, amount: number): number {
  return amount;
}

function multiplyAlpha(current: number, amount: number): number {
  return current * amount;
}

function addAlpha(current: number, amount: number): number {
  return current + amount;
}

const COLOR_TRANSFORMATIONS: readonly ColorTransformation[] = [
  { apply: replaceAlpha, kind: 'alpha', name: 'a:alpha' },
  { apply: multiplyAlpha, kind: 'alpha', name: 'a:alphaMod' },
  { apply: addAlpha, kind: 'alpha', name: 'a:alphaOff' },
  { apply: applyHueMod, kind: 'color', name: 'a:hueMod' },
  { apply: applyLumMod, kind: 'color', name: 'a:lumMod' },
  { apply: applyLumOff, kind: 'color', name: 'a:lumOff' },
  { apply: applySatMod, kind: 'color', name: 'a:satMod' },
  { apply: applyShade, kind: 'color', name: 'a:shade' },
  { apply: applyTint, kind: 'color', name: 'a:tint' },
];

function authoredColorTransformations(
  colorNode: XmlLookupValue,
): AuthoredColorTransformation[] {
  const authored: AuthoredColorTransformation[] = [];
  for (const descriptor of COLOR_TRANSFORMATIONS) {
    for (const transformNode of asArray(nodeAt(colorNode, [descriptor.name]))) {
      const amount = fixedPercentage(textAt(transformNode, ['attrs', 'val']));
      if (amount === undefined) continue;
      authored.push({
        ...descriptor,
        amount,
        order:
          finiteNumber(textAt(transformNode, ['attrs', 'order'])) ??
          Number.MAX_SAFE_INTEGER,
      });
    }
  }
  return authored.sort((left, right) => left.order - right.order);
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

function createMediaData(ref: string): Pick<PptxMediaData, 'blob' | 'ref'> {
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

  const fileExtension = normalizedPath
    .slice(normalizedPath.lastIndexOf('.') + 1)
    .toLowerCase();
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
  const relationshipId = textAt(node, ['a:blip', 'attrs', 'r:embed']);
  if (!relationshipId) return createImageData();

  const relationshipsBySource: Partial<
    Record<string, PptxParserContext['slideResObj'] | undefined>
  > = {
    diagramBg: warpObj.diagramResObj,
    slide: warpObj.slideResObj,
    slideBg: warpObj.slideResObj,
    slideLayoutBg: warpObj.layoutResObj,
    slideMasterBg: warpObj.masterResObj,
    themeBg: warpObj.themeResObj,
  };
  const relationships = relationshipsBySource[source];
  const imagePath = relationships?.[relationshipId]?.target;
  return imagePath ? getImageData(imagePath, warpObj) : createImageData();
}

export function getPicFillOpacity(node: XmlLookupValue): number {
  const amount = textAt(node, ['a:blip', 'a:alphaModFix', 'attrs', 'amt']);
  return clampUnit(fixedPercentage(amount) ?? 1);
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
      const parsedSaturation = fixedPercentage(saturation);
      if (parsedSaturation !== undefined) {
        filters.saturation = parsedSaturation;
      }

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
      const parsedBrightness = fixedPercentage(brightness);
      const parsedContrast = fixedPercentage(contrast);
      if (parsedBrightness !== undefined) {
        filters.brightness = parsedBrightness;
      }
      if (parsedContrast !== undefined) filters.contrast = parsedContrast;

      const sharpenSoften = textAt(effect, [
        'a14:sharpenSoften',
        'attrs',
        'amount',
      ]);
      const amount = fixedPercentage(sharpenSoften);
      if (amount !== undefined) {
        switch (Math.sign(amount)) {
          case 1:
            filters.sharpen = amount;
            break;
          case -1:
            filters.soften = Math.abs(amount);
            break;
        }
      }

      const colorTemperature = textAt(effect, [
        'a14:colorTemperature',
        'attrs',
        'colorTemp',
      ]);
      const parsedColorTemperature = finiteNumber(colorTemperature);
      if (parsedColorTemperature !== undefined) {
        filters.colorTemperature = parsedColorTemperature;
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
  const supported: ReadonlySet<string | undefined> = new Set([
    'circle',
    'rect',
    'shape',
  ]);
  return supported.has(path) ? (path as GradientValue['path']) : 'line';
}

function buildGradient(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
  colorMap?: XmlLookupValue,
  placeholderColor?: string,
): GradientValue {
  const colors = asArray(nodeAt(node, ['a:gsLst', 'a:gs']))
    .map((stop) => {
      const position = clampUnit(
        fixedPercentage(textAt(stop, ['attrs', 'pos'])) ?? 0,
      );
      return {
        position,
        color: getSolidFill(stop, colorMap, placeholderColor, warpObj),
      };
    })
    .sort((left, right) => left.position - right.position)
    .map(({ color, position }) => ({
      color,
      pos: `${position * 100}%`,
    }));

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

  return buildPattern(pattern, warpObj);
}

function buildPattern(
  pattern: XmlLookupValue,
  warpObj: PptxParserContext,
): PatternValue {
  const foregroundColor = getSolidFill(
    nodeAt(pattern, ['a:fgClr']),
    undefined,
    undefined,
    warpObj,
  );
  const backgroundColor = getSolidFill(
    nodeAt(pattern, ['a:bgClr']),
    undefined,
    undefined,
    warpObj,
  );

  return {
    type: textAt(pattern, ['attrs', 'prst']) ?? '',
    foregroundColor: foregroundColor || '#000000',
    backgroundColor: backgroundColor || '#ffffff',
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
  if (placeholderColor === undefined) return null;
  return normalizeHexColor(placeholderColor);
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
  for (const [key, value] of Object.entries(fillList)) {
    if (!key.startsWith('a:')) continue;
    for (const child of asArray(value)) {
      fills.push({
        node: wrapChild(key, child),
        order: finiteNumber(attributes(child).order) ?? Number.MAX_SAFE_INTEGER,
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
  const solid = nodeAt(node, ['a:solidFill']);
  if (solid) {
    const value = getSolidFill(solid, colorMap, placeholderColor, warpObj);
    return value ? { type: 'color', value } : null;
  }
  const gradient = nodeAt(node, ['a:gradFill']);
  if (gradient) {
    return {
      type: 'gradient',
      value: buildGradient(gradient, warpObj, colorMap, placeholderColor),
    };
  }
  const picture = nodeAt(node, ['a:blipFill']);
  if (picture) {
    return {
      type: 'image',
      value: await getBackgroundPicture(node, source, warpObj),
    };
  }
  const pattern = nodeAt(node, ['a:pattFill']);
  if (pattern) {
    return { type: 'pattern', value: buildPattern(pattern, warpObj) };
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
      const themeFill = orderedThemeBackgroundFills(warpObj)[index - 1];
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
  for (const groupNode of [...groupHierarchy].reverse()) {
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
  if (nodeAt(shapeProperties, ['a:noFill'])) return { state: 'none' };

  const solid = nodeAt(shapeProperties, ['a:solidFill']);
  if (solid) {
    const value = getSolidFill(solid, undefined, undefined, warpObj);
    return value
      ? { state: 'found', fill: { type: 'color', value } }
      : { state: 'missing' };
  }
  const gradient = nodeAt(shapeProperties, ['a:gradFill']);
  if (gradient) {
    return {
      state: 'found',
      fill: { type: 'gradient', value: getGradientFill(gradient, warpObj) },
    };
  }
  const pictureNode = nodeAt(shapeProperties, ['a:blipFill']);
  if (pictureNode) {
    const picture = await getPicFill(source, pictureNode, warpObj);
    return {
      state: 'found',
      fill: {
        type: 'image',
        value: { ...picture, opacity: getPicFillOpacity(pictureNode) },
      },
    };
  }
  const pattern = nodeAt(shapeProperties, ['a:pattFill']);
  if (pattern) {
    return {
      state: 'found',
      fill: { type: 'pattern', value: buildPattern(pattern, warpObj) },
    };
  }
  if (nodeAt(shapeProperties, ['a:grpFill'])) {
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

function percentComponent(value: string | undefined): number {
  return clampUnit(fixedPercentage(value) ?? 0);
}

export function getSolidFill(
  solidFill: XmlLookupValue | undefined,
  colorMap: XmlLookupValue | undefined,
  placeholderColor: string | undefined,
  warpObj: PptxParserContext,
): string {
  let colorNode = nodeAt(solidFill, ['a:srgbClr']);
  let color: string;
  if (colorNode) {
    color = normalizedColor(textAt(colorNode, ['attrs', 'val']));
  } else if ((colorNode = nodeAt(solidFill, ['a:schemeClr']))) {
    const scheme = textAt(colorNode, ['attrs', 'val']);
    color =
      scheme === undefined
        ? ''
        : normalizedColor(
            getSchemeColorFromTheme(
              `a:${scheme}`,
              warpObj,
              colorMap,
              placeholderColor,
            ),
          );
  } else if ((colorNode = nodeAt(solidFill, ['a:scrgbClr']))) {
    const values = attributes(colorNode);
    color = normalizedColor(
      byteHex(255 * percentComponent(values.r)) +
        byteHex(255 * percentComponent(values.g)) +
        byteHex(255 * percentComponent(values.b)),
    );
  } else if ((colorNode = nodeAt(solidFill, ['a:prstClr']))) {
    const preset = textAt(colorNode, ['attrs', 'val']);
    color = normalizedColor(
      preset === undefined ? undefined : getColorName2Hex(preset),
    );
  } else if ((colorNode = nodeAt(solidFill, ['a:hslClr']))) {
    const values = attributes(colorNode);
    const hue = finiteNumber(values.hue) ?? 0;
    const rgb = hslToRgb(
      hue / 60_000,
      percentComponent(values.sat),
      percentComponent(values.lum),
    );
    color = normalizedColor(byteHex(rgb.r) + byteHex(rgb.g) + byteHex(rgb.b));
  } else if ((colorNode = nodeAt(solidFill, ['a:sysClr']))) {
    color = normalizedColor(textAt(colorNode, ['attrs', 'lastClr']));
  } else {
    return '';
  }

  if (color === '') return '';

  let alpha: number | undefined;
  for (const transformation of authoredColorTransformations(colorNode)) {
    if (transformation.kind === 'alpha') {
      alpha = clampUnit(
        transformation.apply(alpha ?? 1, transformation.amount),
      );
    } else {
      color = transformation.apply(color, transformation.amount, false);
    }
  }
  if (alpha !== undefined) color = tinycolor(color).setAlpha(alpha).toHex8();
  return color.startsWith('#') ? color : `#${color}`;
}
