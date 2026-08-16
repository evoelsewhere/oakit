import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

function tableFrame(
  id: number,
  columns: readonly string[],
  rows: readonly string[],
): string {
  return `<p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
    <a:graphic><a:graphicData uri="${TABLE_URI}">
      <a:tbl>
        <a:tblPr/>
        <a:tblGrid>${columns.map((width) => `<a:gridCol w="${width}"/>`).join('')}</a:tblGrid>
        ${rows
          .map(
            (height, index) => `<a:tr h="${height}">
              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Row ${index}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
            </a:tr>`,
          )
          .join('')}
      </a:tbl>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`;
}

describe('PowerPoint table dimensions through the public API', () => {
  it('uses canonical positive dimensions and normalizes malformed values', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${tableFrame(820, ['914400', '+457200'], ['457200', '+228600'])}
            ${tableFrame(
              821,
              ['Infinity', '914400junk', '1.5', '-914400', '0'],
              ['Infinity', '457200junk', '1.5', '-457200', '0'],
            )}
          </p:spTree></p:cSld>
        </p:sld>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(byId['820']).toMatchObject({
      colWidths: [72, 36],
      height: 72,
      left: 72,
      rowHeights: [36, 18],
      top: 36,
      type: 'table',
      width: 108,
    });
    expect(byId['821']).toMatchObject({
      colWidths: [0, 0, 0, 0, 0],
      height: 72,
      rowHeights: [0, 0, 0, 0, 0],
      type: 'table',
      width: 144,
    });
    expect(result).not.toHaveProperty('diagnostics');
  });
});
