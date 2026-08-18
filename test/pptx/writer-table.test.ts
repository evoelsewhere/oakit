import { describe, expect, it } from 'vitest';

import type { PptxSceneTableElement } from '../../src/formats/pptx/scene-types';
import { createFieldIdAllocator } from '../../src/formats/pptx/writer/identifiers';
import { serializeTable } from '../../src/formats/pptx/writer/table';

function tableElement(): PptxSceneTableElement {
  return {
    authored: {},
    columns: [100, 200],
    key: 'table-1',
    resolved: { hidden: false },
    rows: [
      {
        cells: [
          {
            fillColor: '#0F172A',
            text: {
              body: { anchor: 'center', wrap: true },
              paragraphs: [
                {
                  children: [
                    {
                      key: 'cell-1-run',
                      properties: {
                        bold: true,
                        color: '#F8FAFC',
                        fontSize: 16,
                      },
                      text: 'Header <&',
                      type: 'run',
                    },
                  ],
                  key: 'cell-1-paragraph',
                },
              ],
            },
          },
          {
            borders: {
              bottom: {
                color: '#38BDF8',
                style: 'dashed',
                width: 1.5,
              },
              left: {
                color: '#F97316',
                style: 'dotted',
                width: 2,
              },
              right: {
                color: '#22C55E',
                style: 'solid',
                width: 2.5,
              },
              top: {
                color: '#A855F7',
                width: 3,
              },
            },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [{ key: 'cell-2-run', text: 'Value', type: 'run' }],
                  key: 'cell-2-paragraph',
                },
              ],
            },
          },
        ],
        height: 40,
      },
      {
        cells: [
          {
            text: {
              body: { anchor: 'bottom' },
              paragraphs: [
                {
                  children: [{ key: 'cell-3-run', text: 'Third', type: 'run' }],
                  key: 'cell-3-paragraph',
                },
              ],
            },
          },
          {
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'cell-4-run', text: 'Fourth', type: 'run' },
                  ],
                  key: 'cell-4-paragraph',
                },
              ],
            },
          },
        ],
        height: 30,
      },
    ],
    type: 'table',
  };
}

describe('native PowerPoint table serialization', () => {
  it('serializes a native graphic frame, grid, rows, and structured text', () => {
    const xml = serializeTable(
      tableElement(),
      { height: 40, width: 300, x: 10, y: 20 },
      2,
      createFieldIdAllocator(),
    );

    expect(xml).toContain(
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>',
    );
    expect(xml).toContain(
      'uri="http://schemas.openxmlformats.org/drawingml/2006/table"',
    );
    expect(xml).toContain(
      '<p:xfrm><a:off x="127000" y="254000"/><a:ext cx="3810000" cy="508000"/></p:xfrm>',
    );
    expect(xml).toContain(
      '<a:tblGrid><a:gridCol w="1270000"/><a:gridCol w="2540000"/></a:tblGrid><a:tr h="508000">',
    );
    expect(xml).toContain('<a:t>Header &lt;&amp;</a:t>');
    expect(xml).toContain(
      '<a:tcPr anchor="ctr"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:tcPr>',
    );
    expect(xml).toContain(
      '<a:lnB w="19050"><a:solidFill><a:srgbClr val="38BDF8"/></a:solidFill><a:prstDash val="dash"/></a:lnB>',
    );
    expect(xml).toContain(
      '<a:lnL w="25400"><a:solidFill><a:srgbClr val="F97316"/></a:solidFill><a:prstDash val="dot"/></a:lnL>',
    );
    expect(xml).toContain(
      '<a:lnR w="31750"><a:solidFill><a:srgbClr val="22C55E"/></a:solidFill><a:prstDash val="solid"/></a:lnR>',
    );
    expect(xml).toContain(
      '<a:lnT w="38100"><a:solidFill><a:srgbClr val="A855F7"/></a:solidFill><a:prstDash val="solid"/></a:lnT>',
    );
    expect(xml).toContain('</a:tr><a:tr h="381000">');
    expect(xml).toContain('<a:tcPr anchor="b"/>');
    expect(xml).toContain('<a:tcPr/>');
    expect(xml).toContain('<a:tc><a:txBody>');
    expect(xml).not.toContain('<a:tc >');
    expect(xml).not.toContain('Stryker was here!');
  });

  it('serializes merged-cell flags, metadata, rotation, and visibility', () => {
    const element = tableElement();
    element.name = `Table <&"'`;
    element.description = `Description <&"'`;
    element.title = `Title <&"'`;
    element.authored.hidden = true;
    const origin = element.rows[0]?.cells[0];
    const continuation = element.rows[0]?.cells[1];
    if (!origin || !continuation) throw new Error('Expected table cells');
    origin.colSpan = 2;
    origin.rowSpan = 2;
    continuation.hMerge = true;
    continuation.vMerge = false;

    const xml = serializeTable(
      element,
      {
        flipHorizontal: false,
        flipVertical: true,
        height: 40,
        rotation: -10,
        width: 300,
        x: 10,
        y: 20,
      },
      7,
      createFieldIdAllocator(),
    );

    expect(xml).toContain(
      '<p:cNvPr id="7" name="Table &lt;&amp;&quot;&apos;" descr="Description &lt;&amp;&quot;&apos;" title="Title &lt;&amp;&quot;&apos;" hidden="1"/>',
    );
    expect(xml).toContain('<p:xfrm rot="-600000" flipH="0" flipV="1">');
    expect(xml).toContain('<a:tc rowSpan="2" gridSpan="2">');
    expect(xml).toContain('<a:tc vMerge="0" hMerge="1">');
  });

  it('allocates fields through the document-scoped writer context', () => {
    const element = tableElement();
    const cell = element.rows[0]?.cells[0];
    if (!cell) throw new Error('Expected table cell');
    cell.text.paragraphs[0] = {
      children: [
        {
          fieldType: 'slidenum',
          key: 'table-field',
          text: '1',
          type: 'field',
        },
      ],
      key: 'field-paragraph',
    };
    const allocation = createFieldIdAllocator();

    expect(
      serializeTable(
        element,
        { height: 40, width: 300, x: 10, y: 20 },
        2,
        allocation,
      ),
    ).toContain('id="{00000000-0000-0000-0000-000000000001}"');
  });

  it.each([
    ['bottom', 'b'],
    ['center', 'ctr'],
    ['distributed', 'dist'],
    ['justified', 'just'],
    ['top', 't'],
  ] as const)('maps table cell anchor %s to %s', (anchor, expected) => {
    const element = tableElement();
    element.columns = [100];
    element.rows = [
      {
        cells: [
          {
            text: {
              body: { anchor },
              paragraphs: [
                {
                  children: [{ key: 'anchor-run', text: '', type: 'run' }],
                  key: 'anchor-paragraph',
                },
              ],
            },
          },
        ],
        height: 40,
      },
    ];

    const xml = serializeTable(
      element,
      { height: 40, width: 100, x: 0, y: 0 },
      2,
      createFieldIdAllocator(),
    );
    expect(xml).toContain(`<a:tcPr anchor="${expected}"/>`);
    expect(xml).not.toContain('Stryker was here!');
  });
});
