import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src/index';
import { createMinimalPptx } from './fixture';

describe('parsePptx', () => {
  it('parses a minimal presentation package', async () => {
    const result = await parsePptx(await createMinimalPptx());

    expect(result.size).toEqual({ width: 720, height: 405 });
    expect(result.themeColors).toEqual([
      '#4472C4',
      '#ED7D31',
      '#A5A5A5',
      '#FFC000',
      '#5B9BD5',
      '#70AD47',
    ]);
    expect(result.usedFonts).toEqual([]);
    expect(result.slides).toHaveLength(1);
    expect(result.slides[0]).toMatchObject({
      fill: { type: 'color', value: '#fff' },
      layoutElements: [],
      note: '',
      transition: null,
    });
    expect(result.slides[0]?.elements).toHaveLength(1);
    expect(result.slides[0]?.elements[0]).toMatchObject({
      id: '2',
      type: 'text',
      left: 72,
      top: 72,
      width: 144,
      height: 72,
      name: 'TextBox 1',
    });
    expect(result.slides[0]?.elements[0]).toHaveProperty(
      'content',
      expect.stringContaining('Hello&nbsp;AI'),
    );
  });
});
