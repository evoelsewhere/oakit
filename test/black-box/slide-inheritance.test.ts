import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  PRESENTATION_NS,
} from './pptx-package';

function layoutWithTransition(transition: string): string {
  return `
    <p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
      ${transition}
    </p:sldLayout>`;
}

function masterWithTransition(transition: string): string {
  return `
    <p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
      <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
      <p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/>
      ${transition}
    </p:sldMaster>`;
}

describe('PowerPoint slide inheritance through the public API', () => {
  it('uses the layout transition before the master transition', async () => {
    const input = await createIndependentPptx({
      'ppt/slideLayouts/slideLayout1.xml': layoutWithTransition(
        '<p:transition spd="fast"><p:push dir="l"/></p:transition>',
      ),
      'ppt/slideMasters/slideMaster1.xml': masterWithTransition(
        '<p:transition spd="med"><p:wipe dir="r"/></p:transition>',
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]?.transition).toEqual({
      type: 'push',
      duration: 500,
      direction: 'l',
    });
  });

  it('uses the master transition when slide and layout omit one', async () => {
    const input = await createIndependentPptx({
      'ppt/slideMasters/slideMaster1.xml': masterWithTransition(
        '<p:transition advClick="0" advTm="321"><p:cover dir="d"/></p:transition>',
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]?.transition).toEqual({
      type: 'cover',
      duration: 1000,
      direction: 'd',
      autoNextAfter: 321,
    });
  });
});
