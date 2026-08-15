import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
} from './pptx-package';

const CONTENT_TYPES_NS =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const STRICT_PRESENTATION_NS = 'http://purl.oclc.org/ooxml/presentationml/main';
const STRICT_OFFICE_REL_NS =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const STRICT_OFFICE_REL_TYPE = `${STRICT_OFFICE_REL_NS}/`;

describe('PPTX package compatibility invariants', () => {
  it('accepts canonical package namespaces behind arbitrary prefixes', async () => {
    const input = await createIndependentPptx({
      '[Content_Types].xml': `
        <ct:Types xmlns:ct="${CONTENT_TYPES_NS}">
          <ct:Override PartName="/ppt/slides/slide1.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>
        </ct:Types>`,
      'ppt/_rels/presentation.xml.rels': `
        <pkg:Relationships xmlns:pkg="${PACKAGE_REL_NS}">
          <pkg:Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="theme/theme1.xml"/>
          <pkg:Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="slides/slide1.xml"/>
        </pkg:Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides).toHaveLength(1);
    expect(result.themeColors).toEqual(['#4472C4']);
  });

  it('recognizes strict presentation namespaces and relationship types', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `
        <strict:presentation xmlns:strict="${STRICT_PRESENTATION_NS}" xmlns:rel="${STRICT_OFFICE_REL_NS}">
          <strict:sldIdLst><strict:sldId id="256" rel:id="rIdSlide1"/></strict:sldIdLst>
          <strict:sldSz cx="9144000" cy="5143500"/>
        </strict:presentation>`,
      'ppt/_rels/presentation.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${STRICT_OFFICE_REL_TYPE}theme" Target="theme/theme1.xml"/>
          <Relationship Id="rIdSlide1" Type="${STRICT_OFFICE_REL_TYPE}slide" Target="slides/slide1.xml"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.size).toEqual({ width: 720, height: 405 });
    expect(result.slides).toHaveLength(1);
  });

  it('ignores relationship ordering when resolving a presentation', async () => {
    const baseline = await createIndependentPptx();
    const reordered = await createIndependentPptx({
      'ppt/_rels/presentation.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="slides/slide1.xml"/>
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="theme/theme1.xml"/>
        </Relationships>`,
    });

    const [expected, actual] = await Promise.all([
      parsePptx(baseline),
      parsePptx(reordered),
    ]);

    expect(actual).toEqual(expected);
  });

  it('resolves internal relationship targets without query or fragment suffixes', async () => {
    const input = await createIndependentPptx({
      'ppt/_rels/presentation.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdTheme" Type="${OFFICE_REL_TYPE}theme" Target="theme/theme1.xml?cache=1"/>
          <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="slides/slide1.xml?revision=2#content"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides).toHaveLength(1);
    expect(result.themeColors).toEqual(['#4472C4']);
  });
});
