import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

interface StyledTextOptions {
  placeholderType?: string;
  runs: string;
  theme?: string;
}

const DEFAULT_THEME = `
  <a:theme xmlns:a="${DRAWING_NS}" name="Typography Theme">
    <a:themeElements>
      <a:clrScheme name="Typography">
        <a:dk1><a:srgbClr val="000000"/></a:dk1>
        <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      </a:clrScheme>
      <a:fontScheme name="Typography">
        <a:majorFont>
          <a:latin typeface="Major Latin"/>
          <a:ea typeface="Major East Asian"/>
          <a:cs typeface="Major Complex"/>
        </a:majorFont>
        <a:minorFont>
          <a:latin typeface="Minor Latin"/>
          <a:ea typeface="Minor East Asian"/>
          <a:cs typeface="Minor Complex"/>
        </a:minorFont>
        <a:fmtScheme name="Typography"/>
      </a:fontScheme>
    </a:themeElements>
  </a:theme>`;

function styledTextSlide({
  placeholderType = 'body',
  runs,
}: StyledTextOptions): string {
  return `
    <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Styled text"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr><p:ph type="${placeholderType}" idx="7"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr lvl="0"/>${runs}</a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`;
}

async function parseStyledText(options: StyledTextOptions): Promise<string> {
  const input = await createIndependentPptx({
    'ppt/slides/slide1.xml': styledTextSlide(options),
    'ppt/theme/theme1.xml': options.theme ?? DEFAULT_THEME,
  });
  const document = await parsePptx(input, { errorMode: 'strict' });
  const element = document.slides[0]?.elements[0];
  expect(element?.type).toBe('text');
  if (element?.type !== 'text') throw new Error('Expected a text element');
  return element.content;
}

function run(properties: string, text: string): string {
  return `<a:r><a:rPr ${properties}/><a:t>${text}</a:t></a:r>`;
}

describe('PowerPoint font styles through the public API', () => {
  it('renders DrawingML boolean and decoration variants', async () => {
    const content = await parseStyledText({
      runs: run('b="true" i="true" u="dbl" strike="dblStrike"', 'Decorated'),
    });

    expect(content).toContain('font-weight: bold;');
    expect(content).toContain('font-style: italic;');
    expect(content).toContain('text-decoration: underline;');
    expect(content).toContain('text-decoration-line: line-through;');
  });

  it('renders exact finite run metrics and shadow CSS', async () => {
    const content = await parseStyledText({
      runs: `<a:r>
        <a:rPr sz="2400" spc="-125" baseline="-25000">
          <a:latin typeface="A Font"/>
          <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
          <a:effectLst>
            <a:outerShdw blurRad="25400" dist="12700" dir="0">
              <a:srgbClr val="445566"/>
            </a:outerShdw>
          </a:effectLst>
        </a:rPr>
        <a:t>Metrics</a:t>
      </a:r>`,
    });

    expect(content).toContain('color: #112233;');
    expect(content).toContain('font-size: 24pt;');
    expect(content).toContain('font-family: &quot;A Font&quot;;');
    expect(content).toContain('letter-spacing: -1.25pt;');
    expect(content).toContain('vertical-align: sub;');
    expect(content).toContain('text-shadow: 1pt 0pt 2pt #445566;');
  });

  it.each([
    ['2400junk', 'Malformed'],
    ['-100', 'Negative'],
    ['0', 'Zero'],
    ['Infinity', 'Infinite'],
  ])('falls back for invalid font size %j', async (size, text) => {
    const content = await parseStyledText({ runs: run(`sz="${size}"`, text) });

    expect(content).toContain('font-size: 18pt;');
    expect(content).not.toContain(
      `font-size: ${Number.parseInt(size) / 100}pt;`,
    );
  });

  it('does not emit partial malformed spacing or baseline values', async () => {
    const content = await parseStyledText({
      runs: run('spc="125junk" baseline="-25000junk"', 'Malformed metrics'),
    });

    expect(content).not.toContain('letter-spacing:');
    expect(content).not.toContain('vertical-align:');
  });

  it('resolves every standard theme font token including complex scripts', async () => {
    const content = await parseStyledText({
      runs: [
        ['a:latin', '+mj-lt', 'MajorLatin'],
        ['a:latin', '+mn-lt', 'MinorLatin'],
        ['a:ea', '+mj-ea', 'MajorEastAsian'],
        ['a:ea', '+mn-ea', 'MinorEastAsian'],
        ['a:cs', '+mj-cs', 'MajorComplex'],
        ['a:cs', '+mn-cs', 'MinorComplex'],
      ]
        .map(
          ([element, typeface, text]) =>
            `<a:r><a:rPr><${element} typeface="${typeface}"/></a:rPr><a:t>${text}</a:t></a:r>`,
        )
        .join(''),
    });

    for (const [family, text] of [
      ['Major Latin', 'MajorLatin'],
      ['Minor Latin', 'MinorLatin'],
      ['Major East Asian', 'MajorEastAsian'],
      ['Minor East Asian', 'MinorEastAsian'],
      ['Major Complex', 'MajorComplex'],
      ['Minor Complex', 'MinorComplex'],
    ]) {
      expect(content).toContain(
        `font-family: &quot;${family}&quot;;">${text}</span>`,
      );
    }
  });

  it.each([
    ['title', 'Major Latin'],
    ['ctrTitle', 'Major Latin'],
    ['subTitle', 'Major Latin'],
    ['body', 'Minor Latin'],
  ])(
    'uses the theme default for %s placeholders',
    async (placeholderType, family) => {
      const content = await parseStyledText({
        placeholderType,
        runs: run('', 'Default theme font'),
      });

      expect(content).toContain(`font-family: &quot;${family}&quot;;`);
    },
  );
});
