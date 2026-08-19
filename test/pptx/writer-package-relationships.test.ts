import { describe, expect, it } from 'vitest';

import {
  serializePresentationRelationships,
  serializeRootRelationships,
  serializeSlideLayoutRelationships,
  serializeSlideMasterRelationships,
  serializeSlideRelationships,
} from '../../src/formats/pptx/writer/package-relationships';

const PREFIX =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
const TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

describe('PowerPoint package relationship graph serialization', () => {
  it('binds the package root to the presentation', () => {
    expect(serializeRootRelationships()).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    );
  });

  it('binds a zero-slide presentation only to its master', () => {
    expect(serializePresentationRelationships(0)).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
    );
  });

  it('binds ordered slides after the master relationship', () => {
    expect(serializePresentationRelationships(3)).toBe(
      `${PREFIX}` +
        `<Relationship Id="rId1" Type="${TYPE}slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rId2" Type="${TYPE}slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rId3" Type="${TYPE}slide" Target="slides/slide2.xml"/>` +
        `<Relationship Id="rId4" Type="${TYPE}slide" Target="slides/slide3.xml"/>` +
        '</Relationships>',
    );
  });

  it('binds the master to its layout before its theme', () => {
    expect(serializeSlideMasterRelationships()).toBe(
      `${PREFIX}` +
        `<Relationship Id="rId1" Type="${TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `<Relationship Id="rId2" Type="${TYPE}theme" Target="../theme/theme1.xml"/>` +
        '</Relationships>',
    );
  });

  it('binds the layout back to its master', () => {
    expect(serializeSlideLayoutRelationships()).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    );
  });

  it('binds every slide to the generated layout', () => {
    expect(serializeSlideRelationships()).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    );
  });

  it('appends deterministic native image relationships after the layout', () => {
    expect(
      serializeSlideRelationships([
        '../media/image1.png',
        '../media/image2.jpeg',
      ]),
    ).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${TYPE}image" Target="../media/image1.png"/><Relationship Id="rId3" Type="${TYPE}image" Target="../media/image2.jpeg"/></Relationships>`,
    );
  });

  it('allocates chart relationships after native images', () => {
    expect(
      serializeSlideRelationships(
        ['../media/image1.png'],
        ['../charts/chart1.xml', '../charts/chart2.xml'],
      ),
    ).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="${TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${TYPE}image" Target="../media/image1.png"/><Relationship Id="rId3" Type="${TYPE}chart" Target="../charts/chart1.xml"/><Relationship Id="rId4" Type="${TYPE}chart" Target="../charts/chart2.xml"/></Relationships>`,
    );
  });

  it.each([-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 10_001])(
    'rejects invalid relationship slide count %s',
    (slideCount) => {
      expect(() => serializePresentationRelationships(slideCount)).toThrow(
        new RangeError(
          'PowerPoint presentation slide count must be an integer from 0 through 10000',
        ),
      );
    },
  );
});
