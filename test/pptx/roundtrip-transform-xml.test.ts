import { describe, expect, it } from 'vitest';

import {
  patchPptxPictureTransformXml,
  patchPptxShapeTransformXml,
} from '../../src/formats/pptx/roundtrip/transform-xml';
import type { PptxRoundTripSetTransformOperation } from '../../src/formats/pptx/roundtrip/types';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MARKUP_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

function slideXml(
  transformAttributes = '',
  offsetAttributes = 'x="254000" y="381000"',
  extentAttributes = 'cx="3810000" cy="1016000"',
  extra = '',
): string {
  return (
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:mc="${MARKUP_NAMESPACE}">` +
    '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr><p:spPr>' +
    `<a:xfrm${transformAttributes}><a:off ${offsetAttributes}/><a:ext ${extentAttributes}/>${extra}</a:xfrm>` +
    '</p:spPr></p:sp></p:spTree></p:cSld></p:sld>'
  );
}

function operation(
  replacement: Partial<PptxRoundTripSetTransformOperation['value']> = {},
  expected: Partial<
    PptxRoundTripSetTransformOperation['expectedTransform']
  > = {},
): PptxRoundTripSetTransformOperation {
  return {
    expectedTransform: {
      flipHorizontal: false,
      flipVertical: false,
      height: 80,
      rotation: 0,
      width: 300,
      x: 20,
      y: 30,
      ...expected,
    },
    id: 'set-transform-1',
    kind: 'set-transform',
    targetKey: 'slide-1-element-1',
    value: {
      flipHorizontal: false,
      flipVertical: false,
      height: 100,
      rotation: 0,
      width: 400,
      x: 50,
      y: 60,
      ...replacement,
    },
  };
}

function pictureXml(): string {
  return slideXml()
    .replace(
      '<p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr>',
      '<p:pic><p:nvPicPr><p:cNvPr id="2"/></p:nvPicPr>',
    )
    .replace('</p:sp></p:spTree>', '</p:pic></p:spTree>');
}

describe('PowerPoint literal shape transform patching', () => {
  it('serializes exact EMUs and omits false optional attributes', () => {
    const input = slideXml();
    const output = patchPptxShapeTransformXml(input, '2', operation());

    expect(output).toBe(
      input.replace(
        '<a:xfrm><a:off x="254000" y="381000"/><a:ext cx="3810000" cy="1016000"/></a:xfrm>',
        '<a:xfrm><a:off x="635000" y="762000"/><a:ext cx="5080000" cy="1270000"/></a:xfrm>',
      ),
    );
    expect(output).not.toContain(' rot=');
    expect(output).not.toContain(' flipH=');
    expect(output).not.toContain(' flipV=');
  });

  it('patches the same strict transform contract on native pictures', () => {
    const input = pictureXml();
    const output = patchPptxPictureTransformXml(input, '2', operation());

    expect(output).toContain(
      '<a:xfrm><a:off x="635000" y="762000"/><a:ext cx="5080000" cy="1270000"/></a:xfrm>',
    );
    expect(output).toContain('<p:pic>');
    expect(output).not.toContain('<p:sp>');
  });

  it.each([
    ['numeric true attributes', ' rot="2700000" flipH="1" flipV="true"'],
    ['lexical false attributes', ' rot="0" flipH="0" flipV="false"'],
  ])('accepts %s', (_name, attributes) => {
    const truthy = attributes.includes('2700000');
    const output = patchPptxShapeTransformXml(
      slideXml(attributes),
      '2',
      operation(
        { flipHorizontal: true, flipVertical: true, rotation: 45 },
        {
          flipHorizontal: truthy,
          flipVertical: truthy,
          rotation: truthy ? 45 : 0,
        },
      ),
    );

    expect(output).toContain('rot="2700000" flipH="1" flipV="1"');
  });

  it.each([
    [
      'x',
      'x="bad" y="381000"',
      'cx="3810000" cy="1016000"',
      'PowerPoint transform x attribute is invalid',
    ],
    [
      'x prefix',
      'x="x254000" y="381000"',
      'cx="3810000" cy="1016000"',
      'PowerPoint transform x attribute is invalid',
    ],
    [
      'x suffix',
      'x="254000x" y="381000"',
      'cx="3810000" cy="1016000"',
      'PowerPoint transform x attribute is invalid',
    ],
    [
      'unsafe x',
      'x="9007199254740992" y="381000"',
      'cx="3810000" cy="1016000"',
      'PowerPoint transform x attribute is unsafe',
    ],
    [
      'cx',
      'x="254000" y="381000"',
      'cx="bad" cy="1016000"',
      'PowerPoint transform cx attribute is invalid',
    ],
  ])('rejects invalid %s attribute', (_name, offset, extent, message) => {
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml('', offset, extent),
        '2',
        operation(),
      ),
    ).toThrow(message);
  });

  it.each(['maybe', '2', 'TRUE'])('rejects invalid flip value %j', (value) => {
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml(` flipH="${value}"`),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint transform flipH attribute is invalid');
  });

  it.each([
    ['PresentationML extension', '<p:extLst/>'],
    ['DrawingML extension', '<a:extLst/>'],
    ['alternate content', '<mc:AlternateContent/>'],
  ])('rejects %s', (_name, markup) => {
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml().replace('</p:spPr>', `${markup}</p:spPr>`),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint transform target contains unsupported compatibility markup',
    );
  });

  it('rejects a source transform that differs from its exact precondition', () => {
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml('', 'x="254001" y="381000"'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint transform source XML does not match its preview precondition',
    );
  });

  it('rejects missing namespaces, target shape, and simple transform', () => {
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml().replace(PRESENTATION_NAMESPACE, 'urn:missing-presentation'),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint text edit slide has no PresentationML namespace');
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml().replace(DRAWING_NAMESPACE, 'urn:missing-drawing'),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint text edit slide has no DrawingML namespace');
    expect(() =>
      patchPptxShapeTransformXml(slideXml(), '7', operation()),
    ).toThrow('PowerPoint text edit requires one unique text shape for id 7');
    expect(() =>
      patchPptxShapeTransformXml(
        slideXml('', undefined, undefined, '<a:chOff x="0" y="0"/>'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint transform target must contain one simple shape transform',
    );
  });
});
