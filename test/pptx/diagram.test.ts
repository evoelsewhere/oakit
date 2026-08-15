import { describe, expect, it, vi } from 'vitest';

import { resolveRelationshipTarget } from '../../src/common';
import type { XmlLookupValue } from '../../src/common';
import type {
  PptxParserContext,
  PptxRelationshipMap,
} from '../../src/formats/pptx/internal/context';
import {
  getDiagramDrawingRelId,
  getDiagramNodeContext,
  getSmartArtTextData,
  loadDiagramFile,
} from '../../src/formats/pptx/internal/diagram';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(
  files: Record<string, XmlLookupValue | null> = {},
  slideResObj: PptxRelationshipMap = {},
) {
  const read = vi.fn((filename: string) =>
    Promise.resolve(files[filename] ?? null),
  );
  const resolve = vi.fn(
    (owner: string, target: string, targetMode?: string) => {
      try {
        return resolveRelationshipTarget(owner, target, targetMode);
      } catch {
        return null;
      }
    },
  );
  const parserContext = {
    diagramFileCache: {},
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    slideResObj,
    themeContent: xml({}),
    xmlReader: { read, resolveRelationshipTarget: resolve },
  } as unknown as PptxParserContext;

  return { context: parserContext, read, resolve };
}

describe('PPTX diagram loading', () => {
  it('does not read an empty part name', async () => {
    const fixture = context();

    await expect(loadDiagramFile(fixture.context, '')).resolves.toBeNull();
    expect(fixture.read).not.toHaveBeenCalled();
  });

  it('caches both present and missing diagram parts', async () => {
    const content = xml({ root: { value: 'present' } });
    const fixture = context({
      'ppt/diagrams/data1.xml': content,
      'ppt/diagrams/missing.xml': null,
    });

    await expect(
      loadDiagramFile(fixture.context, 'ppt/diagrams/data1.xml'),
    ).resolves.toBe(content);
    await loadDiagramFile(fixture.context, 'ppt/diagrams/data1.xml');
    await loadDiagramFile(fixture.context, 'ppt/diagrams/missing.xml');
    await loadDiagramFile(fixture.context, 'ppt/diagrams/missing.xml');

    expect(fixture.read).toHaveBeenCalledTimes(2);
  });

  it('renames drawing element prefixes without rewriting text values', async () => {
    const source = xml({
      'dsp:drawing': {
        'dsp:spTree': { 'a:t': { value: 'literal dsp:value' } },
      },
    });
    const fixture = context({ 'ppt/diagrams/drawing1.xml': source });

    await expect(
      loadDiagramFile(fixture.context, 'ppt/diagrams/drawing1.xml', true),
    ).resolves.toEqual({
      'p:drawing': {
        'p:spTree': { 'a:t': { value: 'literal dsp:value' } },
      },
    });
    expect(source).toEqual({
      'dsp:drawing': {
        'dsp:spTree': { 'a:t': { value: 'literal dsp:value' } },
      },
    });
  });

  it('keeps transformed and original diagram cache entries separate', async () => {
    const source = xml({ 'dsp:drawing': {} });
    const fixture = context({ 'ppt/diagrams/drawing1.xml': source });

    await expect(
      loadDiagramFile(fixture.context, 'ppt/diagrams/drawing1.xml'),
    ).resolves.toBe(source);
    await expect(
      loadDiagramFile(fixture.context, 'ppt/diagrams/drawing1.xml', true),
    ).resolves.toEqual({ 'p:drawing': {} });
    expect(fixture.read).toHaveBeenCalledTimes(2);
  });
});

describe('PPTX diagram semantics', () => {
  it('finds the first valid drawing relationship extension', () => {
    expect(
      getDiagramDrawingRelId(
        xml({
          'dgm:dataModel': {
            'dgm:extLst': {
              'a:ext': [
                { unrelated: {} },
                { 'dsp:dataModelExt': { attrs: { relId: 'rIdDrawing' } } },
                { 'dsp:dataModelExt': { attrs: { relId: 'later' } } },
              ],
            },
          },
        }),
      ),
    ).toBe('rIdDrawing');
    expect(getDiagramDrawingRelId(xml({}))).toBe('');
  });

  it('extracts non-empty SmartArt text across points, paragraphs, and runs', () => {
    expect(
      getSmartArtTextData(
        xml({
          'dgm:dataModel': {
            'dgm:ptLst': {
              'dgm:pt': [
                {
                  'dgm:t': {
                    'a:p': [
                      {
                        'a:r': [
                          { 'a:t': { value: 'Plan' } },
                          { 'a:t': ' now' },
                          { unrelated: {} },
                        ],
                      },
                      { 'a:r': { 'a:t': { value: 'Next' } } },
                    ],
                  },
                },
                { unrelated: {} },
                { 'dgm:t': { 'a:p': { 'a:r': { 'a:t': '' } } } },
              ],
            },
          },
        }),
      ),
    ).toEqual(['Plan now\nNext']);
    expect(getSmartArtTextData(xml({}))).toEqual([]);
  });

  it('resolves a drawing through the diagram data relationship part', async () => {
    const standardPrefix =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
    const strictPrefix =
      'http://purl.oclc.org/ooxml/officeDocument/relationships/';
    const data = xml({
      'dgm:dataModel': {
        'dgm:extLst': {
          'a:ext': {
            'dsp:dataModelExt': { attrs: { relId: 'rIdDrawing' } },
          },
        },
      },
    });
    const drawing = xml({ 'dsp:drawing': { 'dsp:spTree': {} } });
    const fixture = context(
      {
        'ppt/diagrams/data1.xml': data,
        'ppt/diagrams/layout1.xml': xml({ layout: {} }),
        'ppt/diagrams/style1.xml': xml({ style: {} }),
        'ppt/diagrams/colors1.xml': xml({ colors: {} }),
        'ppt/diagrams/_rels/data1.xml.rels': xml({
          Relationships: {
            Relationship: {
              attrs: {
                Id: 'rIdDrawing',
                Target: 'drawing1.xml',
                Type: `${standardPrefix}diagramDrawing`,
              },
            },
          },
        }),
        'ppt/diagrams/drawing1.xml': drawing,
        'ppt/diagrams/_rels/drawing1.xml.rels': xml({
          Relationships: {
            Relationship: [
              {
                attrs: {
                  Id: 'rIdImage',
                  Target: '../media/image1.png',
                  Type: `${standardPrefix}image`,
                },
              },
              {
                attrs: {
                  Id: 'rIdVideo',
                  Target: '../media/video1.mp4',
                  Type: `${strictPrefix}video`,
                },
              },
              {
                attrs: {
                  Id: 'rIdUntyped',
                  Target: '../media/data.bin',
                },
              },
              {
                attrs: {
                  Id: 'unsafe',
                  Target: '../../../outside.png',
                  Type: `${standardPrefix}image`,
                },
              },
              { attrs: { Id: 'missing-target' } },
            ],
          },
        }),
      },
      {
        rIdColors: { target: 'ppt/diagrams/colors1.xml', type: 'colors' },
        rIdData: { target: 'ppt/diagrams/data1.xml', type: 'data' },
        rIdLayout: { target: 'ppt/diagrams/layout1.xml', type: 'layout' },
        rIdStyle: { target: 'ppt/diagrams/style1.xml', type: 'style' },
      },
    );
    const node = xml({
      'a:graphic': {
        'a:graphicData': {
          'dgm:relIds': {
            attrs: {
              'r:cs': 'rIdColors',
              'r:dm': 'rIdData',
              'r:lo': 'rIdLayout',
              'r:qs': 'rIdStyle',
            },
          },
        },
      },
    });

    const result = await getDiagramNodeContext(node, fixture.context);

    expect(result.diagramContent).toMatchObject({
      colors: { colors: {} },
      data,
      layout: { layout: {} },
      quickStyle: { style: {} },
    });
    expect(result.digramFileContent).toEqual({
      'p:drawing': { 'p:spTree': {} },
    });
    expect(result.diagramResObj).toEqual({
      rIdImage: { target: 'ppt/media/image1.png', type: 'image' },
      rIdUntyped: { target: 'ppt/media/data.bin', type: '' },
      rIdVideo: { target: 'ppt/media/video1.mp4', type: 'video' },
    });
    expect(fixture.read).toHaveBeenCalledWith(
      'ppt/diagrams/_rels/data1.xml.rels',
    );
  });
});
