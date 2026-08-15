export type XmlValue =
  XmlNode | XmlValue[] | string | number | boolean | null | undefined;

export interface XmlNode {
  [key: string]: XmlValue;
}

export type XmlReadResult<T> =
  | { status: 'ok'; value: T }
  | { status: 'missing' }
  | { status: 'error'; error: unknown; phase: 'limit' | 'parse' | 'read' };

export interface XmlReadLimits {
  consumeBytes?: (byteLength: number) => void;
  consumeNodes?: (nodeCount: number) => void;
  maxBytes?: number;
  maxDepth?: number;
  maxNodes?: number;
}

export class XmlComplexityLimitError extends Error {
  readonly actual: number;
  readonly limit: number;
  readonly limitName: 'maxTotalXmlNodes' | 'maxXmlDepth' | 'maxXmlNodes';

  constructor(
    limitName: 'maxTotalXmlNodes' | 'maxXmlDepth' | 'maxXmlNodes',
    actual: number,
    limit: number,
  ) {
    super(`XML resource limit ${limitName} exceeded: ${actual} > ${limit}`);
    this.name = 'XmlComplexityLimitError';
    this.actual = actual;
    this.limit = limit;
    this.limitName = limitName;
  }
}

export class XmlStructureError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'XmlStructureError';
  }
}
