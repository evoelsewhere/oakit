import type JSZip from 'jszip';
import { parse } from 'txml/txml';

import {
  readZipEntryBytes,
  ZipExpansionBudgetLimitError,
  ZipEntrySizeLimitError,
} from '../archive/read-entry';
import { simplifyLossless } from './normalize';
import {
  XmlComplexityLimitError,
  type XmlReadLimits,
  type XmlReadResult,
  type XmlValue,
} from './types';
import { assertXmlComplexity, decodeXmlBytes } from './validate';

export { simplifyLossless } from './normalize';
export { getXmlNodeOrder } from './tree';
export { XmlComplexityLimitError, XmlStructureError } from './types';
export type { XmlNode, XmlReadLimits, XmlReadResult, XmlValue } from './types';
export { assertXmlComplexity } from './validate';

/** Read an XML part while preserving missing and failed states for callers. */
export async function readXmlFileResult<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
  limits: XmlReadLimits = {},
): Promise<XmlReadResult<T>> {
  const file = zip.file(filename);
  if (!file) return { status: 'missing' };

  let bytes: Uint8Array;
  try {
    bytes = await readZipEntryBytes(
      file,
      limits.maxBytes ?? Number.MAX_SAFE_INTEGER,
      limits.consumeBytes,
    );
  } catch (error) {
    return {
      status: 'error',
      phase:
        error instanceof ZipEntrySizeLimitError ||
        error instanceof ZipExpansionBudgetLimitError
          ? 'limit'
          : 'read',
      error,
    };
  }

  try {
    const data = decodeXmlBytes(bytes);
    const nodeCount = assertXmlComplexity(data, limits);
    limits.consumeNodes?.(nodeCount);
    return {
      status: 'ok',
      value: simplifyLossless(parse(data, { keepWhitespace: true })) as T,
    };
  } catch (error) {
    return {
      status: 'error',
      phase: error instanceof XmlComplexityLimitError ? 'limit' : 'parse',
      error,
    };
  }
}

/** Read and simplify an XML part. Missing or invalid optional parts return null. */
export async function readXmlFile<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
): Promise<T | null> {
  const result = await readXmlFileResult<T>(zip, filename);
  return result.status === 'ok' ? result.value : null;
}

/** @deprecated Upstream misspelling retained for compatibility. */
export const simplifyLostLess = simplifyLossless;
