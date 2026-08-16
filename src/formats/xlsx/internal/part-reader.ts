import type JSZip from 'jszip';

import {
  readZipEntryBytes,
  ZipExpansionBudgetLimitError,
  ZipEntrySizeLimitError,
} from '../../../common/archive/read-entry';
import {
  readXmlFileResult,
  XmlComplexityLimitError,
  type XmlReadResult,
} from '../../../common/xml/read-xml';
import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxDiagnostic } from '../types';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  resourceLimitDiagnostic,
  XlsxResourceLimitError,
} from './resource-limits';

interface ReadXmlOptions {
  required?: boolean;
}

type ByteLimitName = 'maxMediaBytes' | 'maxPartBytes';

export class XlsxPartReader {
  private readonly xmlCache = new Map<string, XmlReadResult<XmlLookupValue>>();
  private totalExpandedBytes = 0;
  private totalXmlNodes = 0;

  constructor(
    private readonly zip: JSZip,
    private readonly diagnostics: XlsxDiagnostic[],
    private readonly limits: ResolvedXlsxResourceLimits = defaultXlsxResourceLimits(),
  ) {}

  async readXml(
    part: string,
    options: ReadXmlOptions = {},
  ): Promise<XmlLookupValue | null> {
    let result = this.xmlCache.get(part);
    if (!result) {
      result = await readXmlFileResult<XmlLookupValue>(this.zip, part, {
        consumeBytes: (byteLength) => this.consumeExpandedBytes(byteLength),
        consumeNodes: (nodeCount) => this.consumeXmlNodes(nodeCount),
        maxBytes: this.limits.maxXmlBytes,
        maxDepth: this.limits.maxXmlDepth,
        maxNodes: this.limits.maxXmlNodes,
      });
      this.xmlCache.set(part, result);
    }

    if (result.status === 'ok') return result.value;
    if (result.status === 'missing') {
      if (options.required) this.failMissing(part);
      return null;
    }
    return this.failRead(part, result);
  }

  async readBytes(
    part: string,
    limitName: ByteLimitName,
  ): Promise<Uint8Array | null> {
    const entry = this.zip.file(part);
    if (!entry) return null;
    try {
      return await readZipEntryBytes(
        entry,
        this.limits[limitName],
        (byteLength) => this.consumeExpandedBytes(byteLength),
      );
    } catch (error) {
      if (error instanceof ZipEntrySizeLimitError) {
        this.failResource(
          new XlsxResourceLimitError(
            limitName,
            error.actual,
            error.limit,
            part,
          ),
        );
      }
      if (error instanceof ZipExpansionBudgetLimitError) {
        this.failResource(
          new XlsxResourceLimitError(
            'maxTotalUncompressedBytes',
            error.actual,
            error.limit,
            part,
          ),
        );
      }
      throw error;
    }
  }

  private failMissing(part: string): never {
    const diagnostic: XlsxDiagnostic = {
      code: 'missing-required-part',
      message: part
        ? `Required XLSX part is missing: ${part}`
        : 'Required XLSX part name is empty',
      ...(part ? { part } : {}),
      severity: 'error',
    };
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic);
  }

  private failRead(
    part: string,
    result: Extract<XmlReadResult<XmlLookupValue>, { status: 'error' }>,
  ): never {
    if (result.error instanceof ZipEntrySizeLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          'maxXmlBytes',
          result.error.actual,
          result.error.limit,
          part,
        ),
      );
    }
    if (result.error instanceof XmlComplexityLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          result.error.limitName,
          result.error.actual,
          result.error.limit,
          part,
        ),
      );
    }
    if (result.error instanceof ZipExpansionBudgetLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          'maxTotalUncompressedBytes',
          result.error.actual,
          result.error.limit,
          part,
        ),
      );
    }

    const diagnostic: XlsxDiagnostic = {
      code: result.phase === 'parse' ? 'xml-parse-failed' : 'xml-read-failed',
      message: `Failed to ${result.phase} XLSX part ${part}`,
      part,
      severity: 'error',
    };
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause: result.error });
  }

  private consumeExpandedBytes(byteLength: number): void {
    const next = this.totalExpandedBytes + byteLength;
    if (
      !Number.isSafeInteger(next) ||
      next > this.limits.maxTotalUncompressedBytes
    ) {
      throw new ZipExpansionBudgetLimitError(
        next,
        this.limits.maxTotalUncompressedBytes,
      );
    }
    this.totalExpandedBytes = next;
  }

  private consumeXmlNodes(nodeCount: number): void {
    const next = this.totalXmlNodes + nodeCount;
    if (!Number.isSafeInteger(next) || next > this.limits.maxTotalXmlNodes) {
      throw new XmlComplexityLimitError(
        'maxTotalXmlNodes',
        next,
        this.limits.maxTotalXmlNodes,
      );
    }
    this.totalXmlNodes = next;
  }

  private failResource(error: XlsxResourceLimitError): never {
    const diagnostic = resourceLimitDiagnostic(error);
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause: error });
  }
}
