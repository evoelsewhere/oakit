import { describe, expect, it } from 'vitest';

import { svgLinearGradientPaint } from '../../src/formats/pptx/render-svg-gradient';
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
    expect(svgLinearGradientPaint(gradient(), 'pptx-gradient-2-1')).toEqual({
      definition:
        '<linearGradient id="pptx-gradient-2-1" x1="0" y1="0" x2="1" y2="0" gradientTransform="rotate(90 .5 .5)"><stop offset="0%" stop-color="#FFFFFF"/><stop offset="47%" stop-color="#F3F3F3"/><stop offset="100%" stop-color="#434343"/></linearGradient>',
      value: 'url(#pptx-gradient-2-1)',
    });
  });

  it('supports two stops, decimal offsets, and repeated hard-edge offsets', () => {
    expect(
      svgLinearGradientPaint(
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
    [{ path: 'circle' }, 'non-linear path'],
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
      svgLinearGradientPaint(gradient(overrides), 'pptx-gradient-2-1'),
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
    expect(svgLinearGradientPaint(gradient(), id)).toBeNull();
  });

  it('rejects a non-gradient fill', () => {
    expect(
      svgLinearGradientPaint(
        { type: 'color', value: '#FFFFFF' },
        'pptx-gradient-2-1',
      ),
    ).toBeNull();
  });
});
