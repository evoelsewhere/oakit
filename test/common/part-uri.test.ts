import { describe, expect, it } from 'vitest';

import {
  getRelationshipPartUri,
  resolvePartUri,
  resolveRelationshipTarget,
} from '../../src/common/opc/part-uri';

describe('OPC part URI helpers', () => {
  it('builds the relationship part owned by a package part', () => {
    expect(getRelationshipPartUri('ppt/slides/slide1.xml')).toBe(
      'ppt/slides/_rels/slide1.xml.rels',
    );
  });

  it('resolves nested parent segments against the owner directory', () => {
    expect(
      resolvePartUri(
        'ppt/diagrams/drawings/drawing1.xml',
        '../../media/image1.png',
      ),
    ).toBe('ppt/media/image1.png');
  });

  it('normalizes absolute, dot, and Windows-style targets', () => {
    expect(
      resolvePartUri('ppt/slides/slide1.xml', '/ppt/media/image1.png'),
    ).toBe('ppt/media/image1.png');
    expect(resolvePartUri('ppt/slides/slide1.xml', './media/image1.png')).toBe(
      'ppt/slides/media/image1.png',
    );
    expect(
      resolvePartUri('ppt/slides/slide1.xml', '..\\media\\image1.png'),
    ).toBe('ppt/media/image1.png');
  });

  it('preserves targets explicitly marked as external', () => {
    expect(
      resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        'https://example.com/deck?a=1#slide',
        'External',
      ),
    ).toBe('https://example.com/deck?a=1#slide');
  });

  it('rejects unsafe targets that escape the package root', () => {
    expect(() =>
      resolvePartUri('ppt/slides/slide1.xml', '../../../secret.xml'),
    ).toThrow(/package root/);
  });

  it('requires external URIs to declare external target mode', () => {
    expect(() =>
      resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        'https://example.com/image.png',
      ),
    ).toThrow(/TargetMode/);
  });
});
