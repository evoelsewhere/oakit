import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const REGION_STYLE_ID = '{OAKIT-REGIONS}';
const BACKGROUND_STYLE_ID = '{OAKIT-BACKGROUND}';

function cell(label: string, attributes = '', withText = true): string {
  return `<a:tc ${attributes}>
    ${
      withText
        ? `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${label}</a:t></a:r></a:p></a:txBody>`
        : ''
    }
    <a:tcPr/>
  </a:tc>`;
}

function tableFrame(
  id: number,
  attributes: string,
  rows: readonly (readonly string[])[],
  styleId: string | null = REGION_STYLE_ID,
): string {
  const columnCount = Math.max(...rows.map((row) => row.length));
  return `<p:graphicFrame>
    <p:nvGraphicFramePr><p:cNvPr id="${id}" name="Styled table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
    <p:xfrm><a:off x="0" y="0"/><a:ext cx="${columnCount * 914400}" cy="${rows.length * 457200}"/></p:xfrm>
    <a:graphic><a:graphicData uri="${TABLE_URI}">
      <a:tbl>
        <a:tblPr ${attributes}>${
          styleId === null ? '' : `<a:tableStyleId>${styleId}</a:tableStyleId>`
        }</a:tblPr>
        <a:tblGrid>${Array.from(
          { length: columnCount },
          () => '<a:gridCol w="914400"/>',
        ).join('')}</a:tblGrid>
        ${rows
          .map(
            (row) =>
              `<a:tr h="457200">${row.map((value) => cell(value)).join('')}</a:tr>`,
          )
          .join('')}
      </a:tbl>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`;
}

function styleSection(name: string, color: string): string {
  return `<a:${name}>
    <a:tcStyle><a:fill><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:fill></a:tcStyle>
  </a:${name}>`;
}

describe('PowerPoint table style regions through the public API', () => {
  it('maps corners and bands according to every table style flag', async () => {
    const allFlagsRows = Array.from({ length: 3 }, (_, row) =>
      Array.from({ length: 6 }, (_, column) => `R${row}C${column}`),
    );
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${tableFrame(
              830,
              'firstRow="1" firstCol="1" lastRow="1" lastCol="1" bandRow="1" bandCol="1"',
              allFlagsRows,
            )}
            ${tableFrame(
              831,
              'firstRow="0" firstCol="0" lastRow="0" lastCol="0" bandRow="0" bandCol="1"',
              [['C0', 'C1', 'C2', 'C3']],
            )}
            ${tableFrame(
              832,
              'firstRow="1" firstCol="0" lastRow="1" lastCol="0" bandRow="1" bandCol="0"',
              [['R0'], ['R1'], ['R2'], ['R3']],
            )}
            ${tableFrame(
              833,
              'firstRow="0" firstCol="0" lastRow="0" lastCol="0" bandRow="0" bandCol="0"',
              [
                ['A', 'B', 'C', 'D'],
                ['E', 'F', 'G', 'H'],
              ],
            )}
            ${tableFrame(834, '', [['Background']], BACKGROUND_STYLE_ID)}
            ${tableFrame(
              836,
              'firstRow="0" firstCol="1" lastRow="0" lastCol="1" bandRow="0" bandCol="0"',
              [
                ['A', 'B', 'C'],
                ['D', 'E', 'F'],
              ],
            )}
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="835" name="Structural table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></p:xfrm>
              <a:graphic><a:graphicData uri="${TABLE_URI}"><a:tbl>
                <a:tblPr/><a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                <a:tr h="457200">${cell('', 'rowSpan="2" gridSpan="3" vMerge="1" hMerge="1"', false)}</a:tr>
              </a:tbl></a:graphicData></a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/tableStyles.xml': `
        <a:tblStyleLst xmlns:a="${DRAWING_NS}">
          <a:tblStyle styleId="{DECOY}" styleName="Decoy">
            ${styleSection('wholeTbl', 'DECADE')}
          </a:tblStyle>
          <a:tblStyle styleId="${REGION_STYLE_ID}" styleName="Regions">
            ${styleSection('wholeTbl', '010101')}
            ${styleSection('firstRow', '110000')}
            ${styleSection('lastRow', '220000')}
            ${styleSection('band1H', '330000')}
            ${styleSection('band2H', '440000')}
            ${styleSection('firstCol', '001100')}
            ${styleSection('lastCol', '002200')}
            ${styleSection('band1V', '003300')}
            ${styleSection('band2V', '004400')}
            ${styleSection('nwCell', '101010')}
            ${styleSection('neCell', '202020')}
            ${styleSection('swCell', '303030')}
            ${styleSection('seCell', '404040')}
          </a:tblStyle>
          <a:tblStyle styleId="${BACKGROUND_STYLE_ID}" styleName="Background">
            <a:tblBg><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></a:tblBg>
          </a:tblStyle>
        </a:tblStyleLst>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );
    const table = (id: string) => {
      const element = byId[id];
      expect(element?.type).toBe('table');
      if (element?.type !== 'table') throw new Error(`Expected table ${id}`);
      return element;
    };

    expect(
      table('830').data.map((row) => row.map((item) => item.fillColor)),
    ).toEqual([
      ['#101010', '#003300', '#004400', '#003300', '#004400', '#202020'],
      ['#001100', '#003300', '#004400', '#003300', '#004400', '#002200'],
      ['#303030', '#003300', '#004400', '#003300', '#004400', '#404040'],
    ]);
    expect(table('831').data[0]?.map((item) => item.fillColor)).toEqual([
      '#003300',
      '#004400',
      '#003300',
      '#004400',
    ]);
    expect(table('832').data.map((row) => row[0]?.fillColor)).toEqual([
      '#110000',
      '#330000',
      '#440000',
      '#220000',
    ]);
    expect(
      table('833').data.map((row) => row.map((item) => item.fillColor)),
    ).toEqual([
      ['#010101', '#010101', '#010101', '#010101'],
      ['#010101', '#010101', '#010101', '#010101'],
    ]);
    expect(table('834').data[0]?.[0]?.fillColor).toBe('#4472C4');
    expect(
      table('836').data.map((row) => row.map((item) => item.fillColor)),
    ).toEqual([
      ['#001100', '#010101', '#002200'],
      ['#001100', '#010101', '#002200'],
    ]);
    expect(table('835').data[0]?.[0]).toEqual({
      borders: {},
      colSpan: 3,
      hMerge: 1,
      rowSpan: 2,
      text: '',
      vAlign: 'up',
      vMerge: 1,
    });
    expect(table('835').data[0]?.[0]).not.toHaveProperty('fontBold');
  });
});
