import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  OFFICE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

describe('PPTX public API black-box baseline', () => {
  it('parses a standalone spec-shaped package', async () => {
    const input = await createIndependentPptx();

    const result = await parsePptx(input);

    expect(result.size).toEqual({ width: 720, height: 405 });
    expect(result.slides).toHaveLength(1);
    const element = result.slides[0]?.elements[0];
    expect(element?.type).toBe('text');
    if (element?.type !== 'text') throw new Error('Expected a text element');
    expect(element.content).toContain('Black&nbsp;box');
  });

  it('returns unique embedded presentation font names in source order', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
        <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="5143500"/>
        <p:embeddedFontLst>
          <p:embeddedFont><p:font typeface="Aptos"/></p:embeddedFont>
          <p:embeddedFont><p:font typeface="Calibri"/></p:embeddedFont>
          <p:embeddedFont><p:font typeface="Aptos"/></p:embeddedFont>
          <p:embeddedFont><p:font/></p:embeddedFont>
        </p:embeddedFontLst>
      </p:presentation>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.usedFonts).toEqual(['Aptos', 'Calibri']);
  });
});
