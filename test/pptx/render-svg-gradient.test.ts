import { describe, expect, it } from 'vitest';

import { svgGradientPaint } from '../../src/formats/pptx/render-svg-gradient';
import type { Fill } from '../../src/formats/pptx/types';

function gradient(overrides = {}): Fill {
  return {
    type: 'gradient',
    value: {
      colors: [
        { color: '#FFFFFF', pos: '0%' },
        { color: '#F3F3F3', pos: '47%' },
        { color: '#434343', pos: '100%' },
      ],
      path: 'line',
      rot: 90,
      ...overrides,
    },
  };
}

describe('PowerPoint SVG linear gradients', () => {
  it('serializes bounded ordered stops and rotation deterministically', () => {
    expect(svgGradientPaint(gradient(), 'pptx-gradient-2-1')).toEqual({
      definition:
        '<linearGradient id="pptx-gradient-2-1" x1="0" y1="0" x2="1" y2="0" gradientTransform="rotate(90 .5 .5)"><stop offset="0%" stop-color="#FFFFFF"/><stop offset="47%" stop-color="#F3F3F3"/><stop offset="100%" stop-color="#434343"/></linearGradient>',
      value: 'url(#pptx-gradient-2-1)',
    });
  });

  it('supports two stops, decimal offsets, and repeated hard-edge offsets', () => {
    expect(
      svgGradientPaint(
        gradient({
          colors: [
            { color: '#FFFFFF', pos: '47.25%' },
            { color: '#F3F3F3', pos: '47.25%' },
          ],
        }),
        'pptx-gradient-9-8',
      )?.definition,
    ).toContain(
      '<stop offset="47.25%" stop-color="#FFFFFF"/><stop offset="47.25%" stop-color="#F3F3F3"/>',
    );
  });

  it.each([
    [{ path: 'rect' }, 'unsupported rectangular path'],
    [{ path: 'shape' }, 'unsupported shape path'],
    [{ rot: Number.NaN }, 'non-finite rotation'],
    [{ colors: [{ color: '#FFFFFF', pos: '0%' }] }, 'one stop'],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '50%' },
          { color: '#000000', pos: '40%' },
        ],
      },
      'descending stops',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '-1%' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'negative stop',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '0%' },
          { color: '#000000', pos: '101%' },
        ],
      },
      'oversized stop',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '0' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'unitless stop',
    ],
    [
      {
        colors: [
          { color: 'red', pos: '0%' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'unsafe color',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '0%x' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'stop suffix',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: 'x0%' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'stop prefix',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: '47.a%' },
          { color: '#000000', pos: '100%' },
        ],
      },
      'non-numeric decimal',
    ],
    [
      {
        colors: [
          { color: '#FFFFFF', pos: `${'9'.repeat(400)}%` },
          { color: '#000000', pos: '100%' },
        ],
      },
      'non-finite stop',
    ],
  ])('rejects %j (%s)', (overrides, name) => {
    expect(name.length).toBeGreaterThan(0);
    expect(
      svgGradientPaint(gradient(overrides), 'pptx-gradient-2-1'),
    ).toBeNull();
  });

  it.each([
    '',
    'gradient-2-1',
    'pptx-gradient-0-1',
    'pptx-gradient-2-0',
    'x-pptx-gradient-2-1',
    'pptx-gradient-2x-1',
    'pptx-gradient-2-1x',
    'pptx-gradient-2-1" onload="alert(1)',
  ])('rejects unsafe id %j', (id) => {
    expect(svgGradientPaint(gradient(), id)).toBeNull();
  });

  it('rejects a non-gradient fill', () => {
    expect(
      svgGradientPaint(
        { type: 'color', value: '#FFFFFF' },
        'pptx-gradient-2-1',
      ),
    ).toBeNull();
  });

  it('serializes a safe circular gradient as a bounded radial definition', () => {
    expect(
      svgGradientPaint(
        gradient({ path: 'circle', rot: 0 }),
        'pptx-gradient-3-4',
      )?.definition,
    ).toBe(
      '<radialGradient id="pptx-gradient-3-4" cx=".5" cy=".5" r=".5"><stop offset="0%" stop-color="#FFFFFF"/><stop offset="47%" stop-color="#F3F3F3"/><stop offset="100%" stop-color="#434343"/></radialGradient>',
    );
  });
});
