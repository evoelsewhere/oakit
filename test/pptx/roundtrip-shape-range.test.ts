import { describe, expect, it } from 'vitest';

import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditableGraphicFrameXml,
  resolvePptxEditableGroupXml,
  resolvePptxEditablePictureXml,
  resolvePptxEditableShapeXml,
} from '../../src/formats/pptx/roundtrip/shape-range';
import { PptxWriteError } from '../../src/formats/pptx/write-error';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MARKUP_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

function slideXml(shape = '<p:sp><p:cNvPr id="2"/><a:t>Text</a:t></p:sp>') {
  return `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}"><p:cSld><p:spTree>${shape}</p:spTree></p:cSld></p:sld>`;
}

describe('PowerPoint editable shape range', () => {
  it('returns the unique matching shape with canonical prefixes', () => {
    const result = resolvePptxEditableShapeXml(slideXml(), '2');

    expect(result).toMatchObject({
      drawingPrefix: 'a',
      markupPrefix: undefined,
      presentationPrefix: 'p',
      shape: '<p:cNvPr id="2"/><a:t>Text</a:t></p:sp>',
    });
    expect(result.range.end).toBeGreaterThan(result.range.start);
  });

  it('returns a unique native picture without accepting a shape decoy', () => {
    const xml = slideXml(
      '<p:sp><p:cNvPr id="2"/></p:sp><p:pic><p:nvPicPr><p:cNvPr id="2"/></p:nvPicPr><p:spPr/></p:pic>',
    );

    expect(resolvePptxEditablePictureXml(xml, '2').shape).toBe(
      '<p:nvPicPr><p:cNvPr id="2"/></p:nvPicPr><p:spPr/></p:pic>',
    );
    expect(() => resolvePptxEditablePictureXml(slideXml(), '2')).toThrow(
      'PowerPoint text edit requires one unique picture for id 2',
    );
  });

  it('returns a unique native graphic frame without accepting shape ids', () => {
    const xml = slideXml(
      '<p:sp><p:cNvPr id="2"/></p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2"/></p:nvGraphicFramePr><a:graphic/></p:graphicFrame>',
    );

    expect(resolvePptxEditableGraphicFrameXml(xml, '2').shape).toBe(
      '<p:nvGraphicFramePr><p:cNvPr id="2"/></p:nvGraphicFramePr><a:graphic/></p:graphicFrame>',
    );
    expect(() => resolvePptxEditableGraphicFrameXml(slideXml(), '2')).toThrow(
      'PowerPoint text edit requires one unique graphic frame for id 2',
    );
  });

  it('returns a unique native group without accepting nested child ids', () => {
    const xml = slideXml(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2"/></p:nvGrpSpPr><p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4"/></p:nvCxnSpPr></p:cxnSp></p:grpSp>',
    );

    expect(resolvePptxEditableGroupXml(xml, '2').shape).toContain(
      '<p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr></p:sp>',
    );
    expect(() => resolvePptxEditableGroupXml(xml, '3')).toThrow(
      'PowerPoint text edit requires one unique group shape for id 3',
    );
    expect(() => resolvePptxEditableGroupXml(xml, '4')).toThrow(
      'PowerPoint text edit requires one unique group shape for id 4',
    );
  });

  it('supports strict and default presentation namespaces', () => {
    const strict = slideXml()
      .replace(
        PRESENTATION_NAMESPACE,
        'http://purl.oclc.org/ooxml/presentationml/main',
      )
      .replace(DRAWING_NAMESPACE, 'http://purl.oclc.org/ooxml/drawingml/main');
    expect(resolvePptxEditableShapeXml(strict, '2')).toMatchObject({
      drawingPrefix: 'a',
      presentationPrefix: 'p',
    });

    const defaultNamespace = slideXml()
      .replace(
        `xmlns:p="${PRESENTATION_NAMESPACE}"`,
        `xmlns="${PRESENTATION_NAMESPACE}"`,
      )
      .replaceAll('<p:', '<')
      .replaceAll('</p:', '</');
    expect(resolvePptxEditableShapeXml(defaultNamespace, '2')).toMatchObject({
      presentationPrefix: '',
    });
  });

  it('keeps a nested shape id bound to the innermost shape', () => {
    const nested = slideXml(
      '<p:sp><p:cNvPr id="7"/><p:sp><p:cNvPr id="2"/><a:t>Inner</a:t></p:sp><a:t>Outer</a:t></p:sp>',
    );
    const result = resolvePptxEditableShapeXml(nested, '2');

    expect(result.shape).toBe('<p:cNvPr id="2"/><a:t>Inner</a:t></p:sp>');
  });

  it('restores the outer shape after an inner shape closes', () => {
    const nested = slideXml(
      '<p:sp><p:sp><p:cNvPr id="7"/></p:sp><p:cNvPr id="2"/><a:t>Outer</a:t></p:sp>',
    );

    expect(resolvePptxEditableShapeXml(nested, '2').shape).toContain(
      '<p:sp><p:cNvPr id="7"/></p:sp><p:cNvPr id="2"/><a:t>Outer</a:t>',
    );
  });

  it.each([
    ['shape namespace', '<a:sp><p:cNvPr id="2"/></a:sp>'],
    ['shape local name', '<p:other><p:cNvPr id="2"/></p:other>'],
    ['property namespace', '<p:sp><a:cNvPr id="2"/></p:sp>'],
    ['property local name', '<p:sp><p:other id="2"/></p:sp>'],
  ])('ignores a matching id in the wrong %s', (_name, decoy) => {
    const withRealTarget = slideXml(
      `${decoy}<p:sp><p:cNvPr id="2"/><a:t>Real</a:t></p:sp>`,
    );

    expect(resolvePptxEditableShapeXml(withRealTarget, '2').shape).toContain(
      '<a:t>Real</a:t>',
    );
  });

  it('keeps the active shape across ordinary nested elements', () => {
    const nestedProperties = slideXml(
      '<p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr><a:t>Text</a:t></p:sp>',
    );

    expect(resolvePptxEditableShapeXml(nestedProperties, '2').shape).toContain(
      '<p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr>',
    );
  });

  it('returns the markup compatibility prefix', () => {
    const compatible = slideXml().replace(
      '<p:sld ',
      `<p:sld xmlns:mc="${MARKUP_NAMESPACE}" `,
    );

    expect(resolvePptxEditableShapeXml(compatible, '2').markupPrefix).toBe(
      'mc',
    );
  });

  it('retains the XML parser cause for a malformed slide', () => {
    const malformed = `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">`;

    try {
      resolvePptxEditableShapeXml(malformed, '2');
    } catch (error) {
      expect(error).toBeInstanceOf(PptxWriteError);
      if (!(error instanceof PptxWriteError)) throw error;
      expect(error.cause).toBeInstanceOf(Error);
    }
  });

  it.each([
    ['', 'slide root is unsupported'],
    ['<root/>', 'slide root is unsupported'],
    [
      slideXml().replace(PRESENTATION_NAMESPACE, 'urn:missing'),
      'no PresentationML namespace',
    ],
    [
      slideXml().replace(DRAWING_NAMESPACE, 'urn:missing'),
      'no DrawingML namespace',
    ],
    [slideXml(), 'one unique text shape for id 7'],
    [
      slideXml('<p:sp><p:cNvPr id="2"/></p:sp><p:sp><p:cNvPr id="2"/></p:sp>'),
      'one unique text shape for id 2',
    ],
  ])('rejects an unsupported shape document', (xml, message) => {
    expect(() =>
      resolvePptxEditableShapeXml(xml, xml === slideXml() ? '7' : '2'),
    ).toThrow(message);
  });

  it('qualifies names and escapes expression metacharacters', () => {
    expect(qualifiedPptxName('', 't')).toBe('t');
    expect(qualifiedPptxName('a', 't')).toBe('a:t');
    expect(escapePptxXmlPattern('a.dot:t')).toBe('a\\.dot:t');
    expect(pptxShapeHasElement('<a.dot:t/>', 'a.dot:t')).toBe(true);
    expect(pptxShapeHasElement('<aXdot:t/>', 'a.dot:t')).toBe(false);
  });
});
