import type JSZip from 'jszip';

import type { XmlLookupValue } from '../../../common';
import {
  readXmlFileResult,
  type XmlReadResult,
} from '../../../common/xml/read-xml';
import { PptxParseError } from '../errors';
import type {
  PptxDiagnostic,
  PptxErrorMode,
  PptxDiagnosticCode,
} from '../types';

interface ReadPartOptions {
  required?: boolean;
}

function emptyXmlNode(): XmlLookupValue {
  return {} as unknown as XmlLookupValue;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export class PptxXmlReader {
  private readonly cache = new Map<string, XmlReadResult<XmlLookupValue>>();
  private readonly reportedFailures = new Set<string>();

  constructor(
    private readonly zip: JSZip,
    private readonly errorMode: PptxErrorMode,
    private readonly diagnostics: PptxDiagnostic[],
  ) {}

  async read(
    part: string,
    options: ReadPartOptions = {},
  ): Promise<XmlLookupValue> {
    if (!part) {
      if (options.required) this.reportMissing(part);
      return emptyXmlNode();
    }

    let result = this.cache.get(part);
    if (!result) {
      result = await readXmlFileResult<XmlLookupValue>(this.zip, part);
      this.cache.set(part, result);
    }

    if (result.status === 'ok') return result.value;
    if (result.status === 'missing') {
      if (options.required) this.reportMissing(part);
      return emptyXmlNode();
    }

    this.reportReadFailure(part, result, Boolean(options.required));
    return emptyXmlNode();
  }

  private reportMissing(part: string): void {
    const diagnostic: PptxDiagnostic = {
      code: 'missing-required-part',
      message: part
        ? `Required OOXML part is missing: ${part}`
        : 'Required OOXML part name is empty',
      severity: 'error',
      ...(part ? { part } : {}),
    };
    this.report(diagnostic, `missing:${part}`);
  }

  private reportReadFailure(
    part: string,
    result: Extract<XmlReadResult<XmlLookupValue>, { status: 'error' }>,
    required: boolean,
  ): void {
    const code: PptxDiagnosticCode =
      result.phase === 'parse' ? 'xml-parse-failed' : 'xml-read-failed';
    const diagnostic: PptxDiagnostic = {
      code,
      message: `Failed to ${result.phase} OOXML part ${part}: ${errorMessage(
        result.error,
      )}`,
      part,
      severity: required ? 'error' : 'warning',
    };
    this.report(diagnostic, `${code}:${part}`, result.error);
  }

  private report(
    diagnostic: PptxDiagnostic,
    failureKey: string,
    cause?: unknown,
  ): void {
    if (!this.reportedFailures.has(failureKey)) {
      this.reportedFailures.add(failureKey);
      this.diagnostics.push(diagnostic);
    }
    if (this.errorMode === 'strict') {
      throw new PptxParseError(diagnostic, { cause });
    }
  }
}
