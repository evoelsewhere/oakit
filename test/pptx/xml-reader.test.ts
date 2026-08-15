import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxParseError } from '../../src/formats/pptx/errors';
import { PptxXmlReader } from '../../src/formats/pptx/internal/xml-reader';
import type { PptxDiagnostic } from '../../src/formats/pptx/types';

function createArchive(): JSZip {
  const zip = new JSZip();
  zip.file('valid.xml', '<root><child/></root>');
  zip.file('invalid.xml', '<root><child></root>');
  return zip;
}

describe('PptxXmlReader', () => {
  it('reports invalid optional parts and continues in tolerant mode', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await expect(reader.read('invalid.xml')).resolves.toEqual({});
    expect(diagnostics).toMatchObject([
      {
        code: 'xml-parse-failed',
        part: 'invalid.xml',
        severity: 'warning',
      },
    ]);
  });

  it('reports a required missing part once', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await reader.read('missing.xml', { required: true });
    await reader.read('missing.xml', { required: true });

    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      code: 'missing-required-part',
      part: 'missing.xml',
      severity: 'error',
    });
  });

  it('throws a typed error for invalid XML in strict mode', async () => {
    const reader = new PptxXmlReader(createArchive(), 'strict', []);

    await expect(reader.read('invalid.xml')).rejects.toBeInstanceOf(
      PptxParseError,
    );
  });

  it('reports unsafe relationship targets in tolerant mode', () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    expect(
      reader.resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        '../../../secret.xml',
      ),
    ).toBeNull();
    expect(diagnostics).toMatchObject([
      {
        code: 'invalid-relationship-target',
        part: 'ppt/slides/slide1.xml',
        severity: 'warning',
      },
    ]);
  });
});
