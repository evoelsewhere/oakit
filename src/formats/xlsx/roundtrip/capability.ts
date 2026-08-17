import type {
  XlsxCapabilityEntry,
  XlsxCapabilityManifest,
  XlsxEditOperation,
} from './types';

const DOMAINS = [
  'active-content',
  'calculation',
  'cells',
  'charts',
  'comments',
  'conditional-formatting',
  'connections',
  'defined-names',
  'document-properties',
  'drawings-images',
  'external-links',
  'filters-sorts',
  'formulas',
  'hyperlinks',
  'known-extensions',
  'merges',
  'pivots',
  'print-layout',
  'protection',
  'rows-columns',
  'shared-strings',
  'sheet-metadata',
  'sparklines',
  'styles',
  'tables',
  'unknown-extensions',
  'validation',
  'views',
  'workbook-sheets',
] as const;

const OPERATIONS: Array<XlsxEditOperation['kind']> = [
  'add-worksheet',
  'clear-cell',
  'delete-columns',
  'delete-rows',
  'delete-worksheet',
  'insert-columns',
  'insert-rows',
  'rename-worksheet',
  'set-cell',
  'set-cell-style',
  'set-column',
  'set-hyperlink',
  'set-row',
];

export function createXlsxCapabilityManifest(): XlsxCapabilityManifest {
  const domains: XlsxCapabilityEntry[] = DOMAINS.map((domain) => ({
    domain,
    level: 'preservation-only',
  }));
  return {
    domains,
    effectiveLevel: 'R0',
    id: 'xlsx-agent-ready',
    operations: OPERATIONS.map((operation) => ({
      level: 'unsupported',
      operation,
    })),
    producerEvidence: [],
    version: '1',
  };
}
