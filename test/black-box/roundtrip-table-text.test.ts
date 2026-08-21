import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  readPptxRoundTrip,
  renderPptxToSvg,
  replacePptxRoundTripText,
  writePptxRoundTrip,
  type PptxSceneDocument,
} from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

function independentTableSlide(): string {
  const cell = (value: string) =>
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${value}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
  return `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="7" name="Independent table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
              <a:tr h="914400">${cell('Alpha')}${cell('Beta')}</a:tr>
              <a:tr h="914400">${cell('Gamma')}${cell('Delta')}</a:tr>
            </a:tbl>
          </a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`;
}

function independentNestedTextSlide(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="7" name="Independent group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1371600"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="1371600"/></a:xfrm></p:grpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="8" name="Nested text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Nested</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:grpSp>
      </p:spTree></p:cSld>
    </p:sld>`;
}

async function partPayloads(
  data: Uint8Array,
): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const part of Object.values(archive.files)) {
    if (!part.dir) result.set(part.name, await part.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint table cell text editing', () => {
  it('patches one independently authored cell and preserves every other part', async () => {
    const source = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentTableSlide(),
    });
    const snapshot = await readPptxRoundTrip(source);
    const table = snapshot.document.slides[0]?.elements[0];
    expect(table).toMatchObject({ key: 'slide-1-element-1', type: 'table' });
    expect(
      table?.type === 'table'
        ? table.rows.map((row) =>
            row.cells.map((cell) => {
              const child = cell.text.paragraphs[0]?.children[0];
              return child?.type === 'run' ? child.text : undefined;
            }),
          )
        : undefined,
    ).toEqual([
      ['Alpha', 'Beta'],
      ['Gamma', 'Delta'],
    ]);

    const edited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-row-2-cell-1-run-1',
      value: 'Updated <& cell',
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts, verified, rendered] = await Promise.all([
      partPayloads(source),
      partPayloads(output.data),
      readPptxRoundTrip(output.data),
      renderPptxToSvg(output.data, { slideNumbers: [1] }),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(output.report.operations).toMatchObject([
      { kind: 'replace-text', status: 'verified' },
    ]);
    expect(
      verified.document.slides[0]?.elements[0]?.type === 'table'
        ? verified.document.slides[0].elements[0].rows[1]?.cells[0]?.text
            .paragraphs[0]?.children[0]
        : undefined,
    ).toMatchObject({ text: 'Updated <& cell', type: 'run' });

    for (const [name, sourcePayload] of sourceParts) {
      const outputPayload = outputParts.get(name);
      expect(outputPayload, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') {
        expect(outputPayload).not.toEqual(sourcePayload);
        const xml = new TextDecoder().decode(outputPayload);
        expect(xml).toContain(
          '<a:t xml:space="preserve">Updated &lt;&amp; cell</a:t>',
        );
        for (const value of ['Alpha', 'Beta', 'Delta']) {
          expect(xml).toContain(`<a:t>${value}</a:t>`);
        }
      } else {
        expect(outputPayload, name).toEqual(sourcePayload);
      }
    }
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toMatch(/Updated(?:\u00a0| )&lt;&amp;(?:\u00a0| )cell/);
    expect(svg).toContain('Delta');
  });

  it('patches a nested group text run through its complete stable key', async () => {
    const source = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentNestedTextSlide(),
    });
    const snapshot = await readPptxRoundTrip(source);
    expect(snapshot.document.slides[0]?.elements[0]).toMatchObject({
      elements: [
        {
          text: { paragraphs: [{ children: [{ text: 'Nested' }] }] },
          type: 'text',
        },
      ],
      type: 'group',
    });

    const edited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-element-1-run-1',
      value: 'Nested updated',
    });
    const output = await writePptxRoundTrip(edited);
    const verified = await readPptxRoundTrip(output.data);

    expect(output.report).toMatchObject({
      level: 'R2',
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(JSON.stringify(verified.document)).toContain('Nested updated');
  });

  it('patches a table cell nested inside a native group', async () => {
    const scene: PptxSceneDocument = {
      layouts: [],
      masters: [],
      media: [],
      schemaVersion: 2,
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [
            {
              authored: {
                transform: {
                  childSpace: { height: 100, width: 300, x: 0, y: 0 },
                  height: 100,
                  width: 300,
                  x: 72,
                  y: 90,
                },
              },
              elements: [
                {
                  authored: {
                    transform: { height: 100, width: 300, x: 0, y: 0 },
                  },
                  columns: [300],
                  key: 'nested-table',
                  resolved: { hidden: false },
                  rows: [
                    {
                      cells: [
                        {
                          text: {
                            body: {},
                            paragraphs: [
                              {
                                children: [
                                  {
                                    key: 'nested-cell-run',
                                    text: 'Nested cell',
                                    type: 'run',
                                  },
                                ],
                                key: 'nested-cell-paragraph',
                              },
                            ],
                          },
                        },
                      ],
                      height: 100,
                    },
                  ],
                  type: 'table',
                },
              ],
              key: 'table-group',
              resolved: { hidden: false },
              type: 'group',
            },
          ],
          key: 'slide',
        },
      ],
      themes: [],
    };
    const created = await createPptx(scene);
    const snapshot = await readPptxRoundTrip(created.data);
    const edited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-element-1-row-1-cell-1-run-1',
      value: 'Nested table updated',
    });
    const output = await writePptxRoundTrip(edited);

    expect(output.report).toMatchObject({
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(
      JSON.stringify((await readPptxRoundTrip(output.data)).document),
    ).toContain('Nested table updated');
  });

  it('rejects a stale cell precondition without returning a partial package', async () => {
    const source = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentTableSlide(),
    });
    const snapshot = await readPptxRoundTrip(source);
    const edited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-row-1-cell-1-run-1',
      value: 'Updated',
    });
    edited.operations[0] = {
      ...edited.operations[0],
      expectedText: 'Stale',
    } as (typeof edited.operations)[number];

    await expect(writePptxRoundTrip(edited)).rejects.toMatchObject({
      code: 'invalid-snapshot',
      message:
        'PowerPoint round-trip text edit precondition does not match the preview',
    });
  });
});
