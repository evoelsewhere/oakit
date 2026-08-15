import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import { createIndependentPptx } from './pptx-package';

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
});
