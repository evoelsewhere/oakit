import { describe, expect, it } from 'vitest';

import { patchPptxShapeTextXml } from '../../src/formats/pptx/roundtrip/text-xml';
import type { PptxRoundTripReplaceTextOperation } from '../../src/formats/pptx/roundtrip/types';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MARKUP_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const STRICT_PRESENTATION_NAMESPACE =
  'http://purl.oclc.org/ooxml/presentationml/main';
const STRICT_DRAWING_NAMESPACE = 'http://purl.oclc.org/ooxml/drawingml/main';

function slideXml(textNodes: string, extra = ''): string {
  return (
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:mc="${MARKUP_NAMESPACE}">` +
    '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr>' +
    `<p:txBody><a:p><a:r>${textNodes}</a:r></a:p></p:txBody>${extra}` +
    '</p:sp></p:spTree></p:cSld></p:sld>'
  );
}

function operation(expectedText = 'Before'): PptxRoundTripReplaceTextOperation {
  return {
    expectedText,
    id: 'replace-text-1',
    kind: 'replace-text',
    targetKey: 'slide-1-element-1-run-1',
    value: ' After <& ',
  };
}

describe('PowerPoint literal text node patching', () => {
  it.each([
    ['<a:t>Before</a:t>', '<a:t xml:space="preserve">'],
    ['<a:t xml:space="default">Before</a:t>', '<a:t xml:space="preserve">'],
    [
      '<a:t  xml:space = \'preserve\' lang = "en-US">Before</a:t>',
      '<a:t lang = "en-US" xml:space="preserve">',
    ],
    [
      '<a:t lang="en-US" xml:space = "long value">Before</a:t>',
      '<a:t lang="en-US" xml:space="preserve">',
    ],
  ])('replaces text and canonicalizes spacing for %s', (text, opening) => {
    const input = slideXml(text);
    const output = patchPptxShapeTextXml(input, '2', operation());

    expect(output).toBe(
      input.replace(text, `${opening} After &lt;&amp; </a:t>`),
    );
    expect(output.match(/xml:space\s*=/g)).toHaveLength(1);
  });

  it('supports single-quoted and default namespace declarations', () => {
    const singleQuoted = slideXml('<a:t>Before</a:t>')
      .replace(
        `xmlns:p="${PRESENTATION_NAMESPACE}"`,
        `xmlns:p='${PRESENTATION_NAMESPACE}'`,
      )
      .replace(
        `xmlns:a="${DRAWING_NAMESPACE}"`,
        `xmlns:a = '${DRAWING_NAMESPACE}'`,
      );
    expect(patchPptxShapeTextXml(singleQuoted, '2', operation())).toContain(
      '<a:t xml:space="preserve">',
    );

    const defaultNamespace = slideXml('<a:t>Before</a:t>')
      .replace(
        `xmlns:p="${PRESENTATION_NAMESPACE}"`,
        `xmlns="${PRESENTATION_NAMESPACE}"`,
      )
      .replaceAll('<p:', '<')
      .replaceAll('</p:', '</');
    expect(patchPptxShapeTextXml(defaultNamespace, '2', operation())).toContain(
      '<a:t xml:space="preserve">',
    );
    expect(() =>
      patchPptxShapeTextXml(
        defaultNamespace.replace('</sp>', '<extLst/></sp>'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint text edit target contains unsupported compatibility markup',
    );
  });

  it('supports strict PresentationML and DrawingML namespaces', () => {
    const strict = slideXml('<a:t>Before</a:t>')
      .replace(PRESENTATION_NAMESPACE, STRICT_PRESENTATION_NAMESPACE)
      .replace(DRAWING_NAMESPACE, STRICT_DRAWING_NAMESPACE);

    expect(patchPptxShapeTextXml(strict, '2', operation())).toContain(
      '<a:t xml:space="preserve">',
    );
  });

  it('rejects malformed and nested non-slide roots', () => {
    expect(() => patchPptxShapeTextXml('', '2', operation())).toThrow(
      'PowerPoint text edit slide root is unsupported',
    );
    const nestedSlide = `<root><p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}"/></root>`;
    expect(() => patchPptxShapeTextXml(nestedSlide, '2', operation())).toThrow(
      'PowerPoint text edit slide root is unsupported',
    );
    expect(() =>
      patchPptxShapeTextXml(
        `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}">`,
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint text edit slide root is unsupported');
  });

  it('escapes regex metacharacters in namespace prefixes', () => {
    const prefixed = slideXml('<a:t>Before</a:t>')
      .replace(
        `xmlns:a="${DRAWING_NAMESPACE}"`,
        `xmlns:a.dot="${DRAWING_NAMESPACE}" xmlns:aXdot="urn:decoy"`,
      )
      .replaceAll('<a:', '<a.dot:')
      .replaceAll('</a:', '</a.dot:')
      .replace(
        '<a.dot:t>Before</a.dot:t>',
        '<aXdot:t>Decoy</aXdot:t><a.dot:t>Before</a.dot:t>',
      );

    expect(patchPptxShapeTextXml(prefixed, '2', operation())).toContain(
      '<a.dot:t xml:space="preserve">',
    );
  });

  it('ignores misleading ids outside non-visual shape properties', () => {
    const misleading = slideXml('<a:t>Before</a:t>').replace(
      '<p:cNvPr id="2"/>',
      '<p:other id="2"/><p:cNvPr id="7"/>',
    );
    expect(() => patchPptxShapeTextXml(misleading, '2', operation())).toThrow(
      'PowerPoint text edit requires one unique text shape for id 2',
    );
  });

  it('keeps self-closing elements out of the shape stack', () => {
    const withSelfClosing = slideXml('<a:t>Before</a:t>').replace(
      '<p:nvSpPr>',
      '<p:dummy/><p:nvSpPr>',
    );
    expect(patchPptxShapeTextXml(withSelfClosing, '2', operation())).toContain(
      '<a:t xml:space="preserve">',
    );
  });

  it('binds a matching id to the innermost presentation shape', () => {
    const nested = slideXml('<a:t>Outer</a:t>')
      .replace('<p:cNvPr id="2"/>', '<p:cNvPr id="7"/>')
      .replace(
        '<p:txBody>',
        '<p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr>' +
          '<p:txBody><a:p><a:r><a:t>Before</a:t></a:r></a:p></p:txBody>' +
          '</p:sp><p:txBody>',
      );

    const output = patchPptxShapeTextXml(nested, '2', operation());
    expect(output).toContain(
      '<a:t xml:space="preserve"> After &lt;&amp; </a:t>',
    );
    expect(output).toContain('<a:t>Outer</a:t>');
  });

  it.each([
    ['line break', '<a:t>Before</a:t><a:br/>'],
    ['field', '<a:t>Before</a:t><a:fld type="slidenum"/>'],
  ])('rejects a target containing %s', (_name, text) => {
    expect(() =>
      patchPptxShapeTextXml(slideXml(text), '2', operation()),
    ).toThrow('PowerPoint text edit target must contain one plain text run');
  });

  it.each([
    ['PresentationML extension', '<p:extLst/>'],
    ['DrawingML extension', '<a:extLst/>'],
    ['alternate content', '<mc:AlternateContent/>'],
  ])('rejects %s', (_name, markup) => {
    expect(() =>
      patchPptxShapeTextXml(
        slideXml('<a:t>Before</a:t>', markup),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint text edit target contains unsupported compatibility markup',
    );
  });

  it('separately rejects an Office escape and a stale text precondition', () => {
    expect(() =>
      patchPptxShapeTextXml(
        slideXml('<a:t>_x0041_</a:t>'),
        '2',
        operation('_x0041_'),
      ),
    ).toThrow(
      'PowerPoint text edit source XML does not match its preview precondition',
    );
    expect(() =>
      patchPptxShapeTextXml(
        slideXml('<a:t>Before</a:t>'),
        '2',
        operation('Other'),
      ),
    ).toThrow(
      'PowerPoint text edit source XML does not match its preview precondition',
    );
  });

  it.each([
    ['', 'exactly one'],
    ['<a:t>Before</a:t><a:t>Second</a:t>', 'exactly one'],
  ])('rejects text-node cardinality for %j', (text, message) => {
    expect(() =>
      patchPptxShapeTextXml(slideXml(text), '2', operation()),
    ).toThrow(`PowerPoint text edit target must contain ${message} text node`);
  });

  it('rejects missing namespaces and target shape', () => {
    expect(() =>
      patchPptxShapeTextXml(
        slideXml('<a:t>Before</a:t>').replace(
          PRESENTATION_NAMESPACE,
          'urn:missing-presentation',
        ),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint text edit slide has no PresentationML namespace');
    expect(() =>
      patchPptxShapeTextXml(
        slideXml('<a:t>Before</a:t>').replace(
          DRAWING_NAMESPACE,
          'urn:missing-drawing',
        ),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint text edit slide has no DrawingML namespace');
    expect(() =>
      patchPptxShapeTextXml(slideXml('<a:t>Before</a:t>'), '7', operation()),
    ).toThrow('PowerPoint text edit requires one unique text shape for id 7');
  });
});
