import { describe, expect, it } from 'vitest';

import { serializeContentTypes } from '../../src/formats/pptx/writer/content-types';

const PREFIX =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
const FIXED =
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
  '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
  '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
  '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';

describe('PowerPoint content-type serialization', () => {
  it('serializes the fixed hierarchy without inventing a slide', () => {
    expect(serializeContentTypes(0)).toBe(`${PREFIX}${FIXED}</Types>`);
  });

  it('serializes every slide override in numeric order', () => {
    expect(serializeContentTypes(3)).toBe(
      `${PREFIX}${FIXED}` +
        '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
        '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
        '<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
        '</Types>',
    );
  });

  it.each([
    -1,
    1.5,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    Number.MAX_SAFE_INTEGER + 1,
  ])('rejects invalid slide count %s', (slideCount) => {
    expect(() => serializeContentTypes(slideCount)).toThrow(
      new RangeError(
        'PowerPoint slide count must be a non-negative safe integer',
      ),
    );
  });

  it('adds canonical image defaults once for native media', () => {
    const value = serializeContentTypes(1, [
      'image/jpeg',
      'image/png',
      'image/png',
    ]);

    expect(value.match(/Extension="png"/g)).toHaveLength(1);
    expect(value.match(/Extension="jpeg"/g)).toHaveLength(1);
    expect(value).toContain('ContentType="image/png"');
    expect(value).toContain('ContentType="image/jpeg"');
  });

  it('adds every native chart override in numeric order', () => {
    const value = serializeContentTypes(1, [], 2);

    expect(value).toContain(
      '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    );
    expect(value).toContain(
      '<Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    );
    expect(value.indexOf('chart1.xml')).toBeLessThan(
      value.indexOf('chart2.xml'),
    );
  });

  it.each([-1, 1.5, Number.NaN, Number.MAX_SAFE_INTEGER + 1])(
    'rejects invalid chart count %s',
    (chartCount) => {
      expect(() => serializeContentTypes(1, [], chartCount)).toThrow(
        'PowerPoint slide count must be a non-negative safe integer',
      );
    },
  );
});
