import { vi } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type {
  PptxParserContext,
  PptxRelationshipMap,
} from '../../src/formats/pptx/internal/context';
import type { PptxParseOptions } from '../../src/formats/pptx/types';

export function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

interface FillContextOptions {
  layoutRelationships?: PptxRelationshipMap;
  masterRelationships?: PptxRelationshipMap;
  media?: Record<string, Uint8Array | null>;
  options?: PptxParseOptions;
  slideContent?: XmlLookupValue;
  slideLayoutContent?: XmlLookupValue;
  slideMasterContent?: XmlLookupValue;
  slideRelationships?: PptxRelationshipMap;
  themeContent?: XmlLookupValue;
  themeRelationships?: PptxRelationshipMap;
}

export function fillContext(values: FillContextOptions = {}) {
  const media = values.media ?? {};
  const readMedia = vi.fn((path: string) =>
    Promise.resolve(media[path] ?? null),
  );
  const context = {
    diagramFileCache: {},
    diagramResObj: {},
    layoutResObj: values.layoutRelationships ?? {},
    loadedAudios: {},
    loadedImages: {},
    loadedVideos: {},
    masterResObj: values.masterRelationships ?? {},
    options: {
      audioMode: 'none',
      errorMode: 'tolerant',
      imageMode: 'base64',
      videoMode: 'none',
      ...values.options,
    },
    slideContent: values.slideContent ?? xml({}),
    slideLayoutContent: values.slideLayoutContent ?? xml({}),
    slideMasterContent: values.slideMasterContent ?? xml({}),
    slideResObj: values.slideRelationships ?? {},
    themeContent: values.themeContent ?? xml({}),
    themeResObj: values.themeRelationships ?? {},
    xmlReader: { readMedia },
  } as unknown as PptxParserContext;

  return { context, readMedia };
}
