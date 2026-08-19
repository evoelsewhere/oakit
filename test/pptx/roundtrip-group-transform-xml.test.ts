import { describe, expect, it } from 'vitest';

import type { PptxRoundTripSetTransformOperation } from '../../src/formats/pptx/roundtrip/types';
import { patchPptxGroupTransformXml } from '../../src/formats/pptx/roundtrip/transform-xml';

const PRESENTATION =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MARKUP = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function xml(extra = '', transform = ''): string {
  return `<p:sld xmlns:p="${PRESENTATION}" xmlns:a="${DRAWING}" xmlns:mc="${MARKUP}"><p:cSld><p:spTree><p:grpSp><p:nvGrpSpPr><p:cNvPr id="2"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm${transform}><a:off x="254000" y="381000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm>${extra}</p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr></p:sp></p:grpSp></p:spTree></p:cSld></p:sld>`;
}

function operation(
  replacement: Partial<
    Extract<
      PptxRoundTripSetTransformOperation['value'],
      { childSpace: unknown }
    >
  > = {},
  expected: Partial<
    Extract<
      PptxRoundTripSetTransformOperation['expectedTransform'],
      { childSpace: unknown }
    >
  > = {},
): PptxRoundTripSetTransformOperation {
  return {
    expectedTransform: {
      childSpace: { height: 100, width: 100, x: 10, y: 20 },
      flipHorizontal: false,
      flipVertical: false,
      height: 100,
      rotation: 0,
      width: 200,
      x: 20,
      y: 30,
      ...expected,
    },
    id: 'set-transform-1',
    kind: 'set-transform',
    targetKey: 'slide-1-element-1',
    value: {
      childSpace: { height: 120, width: 150, x: 0, y: 0 },
      flipHorizontal: true,
      flipVertical: true,
      height: 180,
      rotation: 15,
      width: 300,
      x: 40,
      y: 50,
      ...replacement,
    },
  };
}

describe('PowerPoint group transform patching', () => {
  it('patches outer and child coordinate spaces with exact EMUs', () => {
    const input = xml();
    const sourceTransform =
      '<a:xfrm><a:off x="254000" y="381000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm>';
    const replacementTransform =
      '<a:xfrm rot="900000" flipH="1" flipV="1"><a:off x="508000" y="635000"/><a:ext cx="3810000" cy="2286000"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="1524000"/></a:xfrm>';
    const output = patchPptxGroupTransformXml(input, '2', operation());

    expect(output).toBe(input.replace(sourceTransform, replacementTransform));
  });

  it('binds source rotation and flip attributes exactly', () => {
    const output = patchPptxGroupTransformXml(
      xml('', ' rot="900000" flipH="1" flipV="1"'),
      '2',
      operation({}, { flipHorizontal: true, flipVertical: true, rotation: 15 }),
    );

    expect(output).toContain('<a:xfrm rot="900000" flipH="1" flipV="1">');
  });

  it('treats omitted expected optional attributes as false and zero', () => {
    const omitted = operation();
    delete omitted.expectedTransform.flipHorizontal;
    delete omitted.expectedTransform.flipVertical;
    delete omitted.expectedTransform.rotation;

    expect(() =>
      patchPptxGroupTransformXml(xml(), '2', omitted),
    ).not.toThrow();
  });

  it('supports namespace aliases and omits false optional attributes', () => {
    const aliased = xml()
      .replace('xmlns:p=', 'xmlns:p1=')
      .replace('xmlns:a=', 'xmlns:d=')
      .replaceAll('<p:', '<p1:')
      .replaceAll('</p:', '</p1:')
      .replaceAll('<a:', '<d:')
      .replaceAll('</a:', '</d:');
    const output = patchPptxGroupTransformXml(
      aliased,
      '2',
      operation({ flipHorizontal: false, flipVertical: false, rotation: 0 }),
    );

    expect(output).toContain('<d:xfrm><d:off');
    expect(output).not.toContain(' flipH=');
    expect(output).not.toContain(' flipV=');
    expect(output).not.toContain(' rot=');
  });

  it('rejects a stale outer or child-space precondition', () => {
    expect(() =>
      patchPptxGroupTransformXml(
        xml().replace('x="127000" y="254000"', 'x="127001" y="254000"'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint group source XML does not match its preview precondition',
    );
    expect(() =>
      patchPptxGroupTransformXml(
        xml().replace('cx="2540000"', 'cx="2540001"'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint group source XML does not match its preview precondition',
    );
  });

  it.each(['<p:extLst/>', '<a:extLst/>', '<mc:AlternateContent/>'])(
    'rejects compatibility markup %s',
    (markup) => {
      expect(() =>
        patchPptxGroupTransformXml(xml(markup), '2', operation()),
      ).toThrow(
        'PowerPoint group transform target contains unsupported compatibility markup',
      );
    },
  );

  it('rejects missing targets and non-simple group transforms', () => {
    expect(() => patchPptxGroupTransformXml(xml(), '7', operation())).toThrow(
      'PowerPoint text edit requires one unique group shape for id 7',
    );
    expect(() =>
      patchPptxGroupTransformXml(
        xml().replace('<a:chExt', '<a:extra/><a:chExt'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint group target must contain one simple group transform',
    );
    expect(() =>
      patchPptxGroupTransformXml(
        xml().replaceAll('p:grpSpPr', 'p:otherGroupProperties'),
        '2',
        operation(),
      ),
    ).toThrow(
      'PowerPoint group target must contain one direct group property block',
    );
  });

  it('rejects operations without group child spaces', () => {
    const missingExpected = operation();
    missingExpected.expectedTransform = {
      height: 100,
      width: 200,
      x: 20,
      y: 30,
    };
    expect(() =>
      patchPptxGroupTransformXml(xml(), '2', missingExpected),
    ).toThrow('PowerPoint group transform has no child space');

    const missingValue = operation();
    missingValue.value = { height: 100, width: 200, x: 20, y: 30 };
    expect(() => patchPptxGroupTransformXml(xml(), '2', missingValue)).toThrow(
      'PowerPoint group transform has no child space',
    );
  });
});
