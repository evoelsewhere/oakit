import { describe, expect, it } from 'vitest';

import { findUnsafeSvgFeatures } from './svg-safety';

const SVG = 'http://www.w3.org/2000/svg';
const XLINK = 'http://www.w3.org/1999/xlink';

function svg(body: string): string {
  return `<svg xmlns="${SVG}" xmlns:xlink="${XLINK}">${body}</svg>`;
}

describe('corpus SVG safety oracle', () => {
  it('does not mistake an authored URL for a resource reference', () => {
    expect(
      findUnsafeSvgFeatures(
        svg(
          '<text>Visit https://example.test or type url(https://example.test)</text>',
        ),
      ),
    ).toEqual([]);
  });

  it.each([
    ['HTTPS image', '<image href="https://attacker.test/pixel.png"/>'],
    ['file image', '<image href="file:///etc/passwd"/>'],
    ['blob image', '<image href="blob:https://attacker.test/id"/>'],
    ['relative image', '<image href="neighbor.png"/>'],
    [
      'external use',
      '<use href="#local" xlink:href="https://attacker.test/a"/>',
    ],
    ['inline CSS URL', '<rect style="fill:url(https://attacker.test/a)"/>'],
    ['stylesheet import', '<style>@import "https://attacker.test/a";</style>'],
    [
      'XML stylesheet',
      '<?xml-stylesheet href="https://attacker.test/a"?><rect/>',
    ],
  ])('reports an unsafe %s reference', (_label, body) => {
    expect(findUnsafeSvgFeatures(svg(body))).not.toEqual([]);
  });

  it.each([
    ['script', '<script>alert(1)</script>'],
    [
      'foreignObject',
      '<foreignObject><iframe xmlns="http://www.w3.org/1999/xhtml"/></foreignObject>',
    ],
  ])('reports the forbidden %s element', (_label, body) => {
    expect(findUnsafeSvgFeatures(svg(body))).not.toEqual([]);
  });

  it.each([
    ['fragment reference', '<use href="#shape"/>'],
    ['fragment CSS URL', '<rect fill="url(#gradient)"/>'],
    ['PNG data URL', '<image href="data:image/png;base64,iVBORw0KGgo="/>'],
    ['JPEG data URL', '<image href="data:image/jpeg;base64,/9j/2Q=="/>'],
  ])('accepts a self-contained %s', (_label, body) => {
    expect(findUnsafeSvgFeatures(svg(body))).toEqual([]);
  });
});
