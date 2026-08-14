import type JSZip from 'jszip';

import type { XmlLookupValue } from '../../../common';
import type { PptxParseOptions } from '../types';

export interface PptxRelationship {
  target: string;
  type: string;
}

export type PptxRelationshipMap = Record<string, PptxRelationship>;

export interface PptxDiagramContent {
  colors: XmlLookupValue | null;
  data: XmlLookupValue | null;
  drawing: XmlLookupValue | null;
  layout: XmlLookupValue | null;
  quickStyle: XmlLookupValue | null;
}

export interface PptxMediaData {
  base64: string;
  blob: string;
  ref: string;
}

export interface PptxNodeIndex {
  idTable: Record<string, XmlLookupValue>;
  idxTable: Record<string, XmlLookupValue>;
  typeTable: Record<string, XmlLookupValue>;
}

export interface PptxParserContext {
  defaultTextStyle: XmlLookupValue;
  diagramContent?: PptxDiagramContent;
  diagramFileCache: Record<string, XmlLookupValue | null>;
  diagramResObj?: PptxRelationshipMap;
  digramFileContent?: XmlLookupValue;
  layoutResObj: PptxRelationshipMap;
  loadedAudios: Record<string, PptxMediaData>;
  loadedImages: Record<string, PptxMediaData>;
  loadedVideos: Record<string, PptxMediaData>;
  masterResObj: PptxRelationshipMap;
  options: Required<PptxParseOptions>;
  slideContent: XmlLookupValue;
  slideLayoutContent: XmlLookupValue;
  slideLayoutTables: PptxNodeIndex;
  slideMasterContent: XmlLookupValue;
  slideMasterTables: PptxNodeIndex;
  slideMasterTextStyles: XmlLookupValue;
  slideResObj: PptxRelationshipMap;
  tableStyles: XmlLookupValue;
  themeContent: XmlLookupValue;
  themeResObj: PptxRelationshipMap;
  zip: JSZip;
}
