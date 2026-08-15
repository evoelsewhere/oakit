import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

function paragraphSlide(
  bodyAttributes: string,
  bodyChildren: string,
  paragraphProperties: string,
) {
  return `
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Paragraph"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr><p:ph type="body" idx="7"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr ${bodyAttributes}>${bodyChildren}</a:bodyPr><a:lstStyle/>
            <a:p>
              <a:pPr ${paragraphProperties}>
                <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
                <a:spcBef><a:spcPct val="50000"/></a:spcBef>
                <a:spcAft><a:spcPts val="1200"/></a:spcAft>
              </a:pPr>
              <a:r><a:rPr/><a:t>Paragraph metrics</a:t></a:r>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`;
}

async function parseParagraph(
  bodyAttributes: string,
  bodyChildren: string,
  paragraphProperties: string,
) {
  const document = await parsePptx(
    await createIndependentPptx({
      'ppt/slides/slide1.xml': paragraphSlide(
        bodyAttributes,
        bodyChildren,
        paragraphProperties,
      ),
    }),
    { errorMode: 'strict' },
  );
  const element = document.slides[0]?.elements[0];
  expect(element?.type).toBe('text');
  if (element?.type !== 'text') throw new Error('Expected a text element');
  return element;
}

describe('PowerPoint paragraph styles through the public API', () => {
  it('renders exact alignment, spacing, indentation, and autofit semantics', async () => {
    const element = await parseParagraph(
      'anchor="ctr"',
      '<a:normAutofit fontScale="75000"/>',
      'lvl="0" algn="justLow" marL="12700" indent="-6350"',
    );

    expect(element).toMatchObject({
      autoFit: { type: 'text', fontScale: 75 },
      vAlign: 'mid',
    });
    expect(element.content).toContain('text-align: justify;');
    expect(element.content).toContain('line-height: 1.5;');
    expect(element.content).toContain('margin-top: 0.5em;');
    expect(element.content).toContain('margin-bottom: 12pt;');
    expect(element.content).toContain('margin-left: 1pt;');
    expect(element.content).toContain('text-indent: -0.5pt;');
  });

  it('does not emit non-finite CSS or partial malformed metrics', async () => {
    const element = await parseParagraph(
      'anchor="invalid"',
      '<a:normAutofit fontScale="75000junk"/>',
      'lvl="0junk" algn="thaiDist" marL="12700junk" indent="Infinity"',
    );

    expect(element.vAlign).toBe('up');
    expect(element.autoFit).toEqual({ type: 'text' });
    expect(element.content).toContain('text-align: justify;');
    expect(element.content).not.toContain('margin-left:');
    expect(element.content).not.toContain('text-indent:');
    expect(element.content).not.toMatch(/NaN|Infinity/);
  });
});
