import JSZip from 'jszip';

import { XlsxParseError } from './errors';
import {
  assertXlsxArchiveWithinLimits,
  assertXlsxInputWithinLimits,
  copyXlsxInputBytes,
} from './internal/archive';
import { XlsxPartReader } from './internal/part-reader';
import {
  resolveXlsxResourceLimits,
  resourceLimitDiagnostic,
  XlsxResourceLimitError,
} from './internal/resource-limits';
import { discoverXlsxWorkbook } from './internal/workbook-discovery';
import { parseXlsxWorkbookManifest } from './internal/workbook-manifest';
import type {
  XlsxDiagnostic,
  XlsxDocument,
  XlsxInput,
  XlsxParseOptions,
  XlsxParseResult,
} from './types';

function failResource(
  error: XlsxResourceLimitError,
  diagnostics: XlsxDiagnostic[],
): never {
  const diagnostic = resourceLimitDiagnostic(error);
  diagnostics.push(diagnostic);
  throw new XlsxParseError(diagnostic, { cause: error });
}

function assertOptions(options: XlsxParseOptions): void {
  if (
    options.errorMode !== undefined &&
    options.errorMode !== 'strict' &&
    options.errorMode !== 'tolerant'
  ) {
    throw new TypeError('XLSX errorMode is invalid');
  }
  if (
    options.displayTextMode !== undefined &&
    options.displayTextMode !== 'none' &&
    options.displayTextMode !== 'supported'
  ) {
    throw new TypeError('XLSX displayTextMode is invalid');
  }
  if (
    options.imageMode !== undefined &&
    options.imageMode !== 'base64' &&
    options.imageMode !== 'blob' &&
    options.imageMode !== 'both' &&
    options.imageMode !== 'none'
  ) {
    throw new TypeError('XLSX imageMode is invalid');
  }
  if (
    options.pivotCacheMode !== undefined &&
    options.pivotCacheMode !== 'metadata' &&
    options.pivotCacheMode !== 'none' &&
    options.pivotCacheMode !== 'records'
  ) {
    throw new TypeError('XLSX pivotCacheMode is invalid');
  }
}

async function openXlsxPackage(
  input: XlsxInput,
  diagnostics: XlsxDiagnostic[],
  limits: ReturnType<typeof resolveXlsxResourceLimits>,
): Promise<JSZip> {
  try {
    assertXlsxInputWithinLimits(input, limits);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await copyXlsxInputBytes(input));
  } catch (cause) {
    const diagnostic: XlsxDiagnostic = {
      code: 'invalid-package',
      message: 'Failed to open XLSX OPC package',
      severity: 'error',
    };
    diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause });
  }

  try {
    assertXlsxArchiveWithinLimits(zip, limits);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }
  return zip;
}

export async function parseXlsxWithDiagnostics(
  input: XlsxInput,
  options: XlsxParseOptions = {},
): Promise<XlsxParseResult> {
  assertOptions(options);
  const diagnostics: XlsxDiagnostic[] = [];
  const limits = resolveXlsxResourceLimits(options.limits);
  const zip = await openXlsxPackage(input, diagnostics, limits);
  const reader = new XlsxPartReader(zip, diagnostics, limits);
  const discovery = await discoverXlsxWorkbook(reader, limits);
  const manifest = await parseXlsxWorkbookManifest(discovery, reader, limits);
  const document: XlsxDocument = {
    sheets: manifest.sheets,
    styles: [],
    workbook: manifest.properties,
  };
  return { diagnostics, document };
}

export async function parseXlsx(
  input: XlsxInput,
  options: XlsxParseOptions = {},
): Promise<XlsxDocument> {
  return (await parseXlsxWithDiagnostics(input, options)).document;
}
