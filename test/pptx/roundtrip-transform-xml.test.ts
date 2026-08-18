import { describe, expect, it } from 'vitest';

import {
  patchPptxGraphicFrameTransformXml,
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

function tableXml(
  columns = ['1270000', '2540000'],
  rows = ['381000', '635000'],
): string {
  return (
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:mc="${MARKUP_NAMESPACE}">` +
    '<p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2"/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="254000" y="381000"/><a:ext cx="3810000" cy="1016000"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>' +
    `<a:tblGrid>${columns.map((width) => `<a:gridCol w="${width}"/>`).join('')}</a:tblGrid>` +
    rows
      .map(
        (height) =>
          `<a:tr h="${height}"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>`,
      )
      .join('') +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>'
  );
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

  it('patches a native table frame and proportionally resizes its exact grid', () => {
    const output = patchPptxGraphicFrameTransformXml(
      tableXml(),
      '2',
      operation(),
    );

    expect(output).toContain(
      '<p:xfrm><a:off x="635000" y="762000"/><a:ext cx="5080000" cy="1270000"/></p:xfrm>',
    );
    expect(output).toContain(
      '<a:gridCol w="1693333"/><a:gridCol w="3386667"/>',
    );
    expect(output).toContain('<a:tr h="476250">');
    expect(output).toContain('<a:tr h="793750">');
    expect(output).toContain('<p:graphicFrame>');
    expect(output.match(/<p:sld\b/g)).toHaveLength(1);
    expect(output.match(/<\/p:sld>/g)).toHaveLength(1);
  });

  it('distributes integer rounding while preserving every tiny grid unit', () => {
    const input = tableXml(['1', '1', '1', '1'], ['1'])
      .replace('cx="3810000"', 'cx="4"')
      .replace('cy="1016000"', 'cy="1"');
    const exactBoundary = patchPptxGraphicFrameTransformXml(
      input,
      '2',
      operation(
        { height: 1 / 12_700, width: 4 / 12_700 },
        { height: 1 / 12_700, width: 4 / 12_700 },
      ),
    );
    expect(exactBoundary.match(/<a:gridCol w="1"\/>/g)).toHaveLength(4);

    const rounded = patchPptxGraphicFrameTransformXml(
      input,
      '2',
      operation(
        { height: 1 / 12_700, width: 6 / 12_700 },
        { height: 1 / 12_700, width: 4 / 12_700 },
      ),
    );
    expect(rounded).toContain(
      '<a:gridCol w="2"/><a:gridCol w="2"/><a:gridCol w="1"/><a:gridCol w="1"/>',
    );
    expect(rounded.match(/<a:gridCol\b/g)).toHaveLength(4);

    const lastRemainderInput = tableXml(['1', '1', '1'], ['1'])
      .replace('cx="3810000"', 'cx="3"')
      .replace('cy="1016000"', 'cy="1"');
    const lastRemainder = patchPptxGraphicFrameTransformXml(
      lastRemainderInput,
      '2',
      operation(
        { height: 1 / 12_700, width: 4 / 12_700 },
        { height: 1 / 12_700, width: 3 / 12_700 },
      ),
    );
    expect(lastRemainder).toContain(
      '<a:gridCol w="1"/><a:gridCol w="1"/><a:gridCol w="2"/>',
    );
  });

  it('rejects table grids that disagree with the semantic precondition', () => {
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml(['1270001', '2540000']),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint table column widths do not match the preview precondition',
    );
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml(undefined, ['381001', '634999']),
        '2',
        operation(),
      ),
    ).not.toThrow();
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml(undefined, ['381001', '635000']),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint table row heights do not match the preview precondition',
    );
  });

  it('rejects missing, non-positive, and unrepresentable table dimensions', () => {
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml([], ['381000', '635000']),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint table has no column widths');
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml(['0', '3810000']),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint table column widths do not match the preview precondition',
    );
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml(),
        '2',
        operation({ width: 0.0001 }),
      ),
    ).toThrow(
      'PowerPoint table column widths cannot fit the requested transform',
    );
    expect(() =>
      patchPptxGraphicFrameTransformXml(
        tableXml().replaceAll('a:tbl', 'a:notTbl'),
        '2',
        operation(),
      ),
    ).toThrow('PowerPoint graphic frame is not a native table');
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
