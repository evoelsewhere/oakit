import { describe, expect, it, vi } from 'vitest';

import { eachElement, getTextByPathList } from '../../src/common/xml/tree';

describe('compatibility XML tree traversal', () => {
  it('returns nested scalar, object, and array values', () => {
    const tree = {
      root: {
        attrs: { id: '7' },
        children: [{ value: 'first' }, { value: 'second' }],
      },
    };

    expect(getTextByPathList<string>(tree, ['root', 'attrs', 'id'])).toBe('7');
    expect(getTextByPathList(tree, ['root', 'attrs'])).toEqual({ id: '7' });
    expect(getTextByPathList(tree, ['root', 'children'])).toHaveLength(2);
  });

  it.each([
    [null, ['root']],
    [undefined, ['root']],
    ['text', ['root']],
    [{ root: null }, ['root', 'child']],
    [{ root: undefined }, ['root', 'child']],
    [{ root: {} }, ['root', 'missing']],
  ])('returns undefined when path %j cannot be traversed', (node, path) => {
    expect(getTextByPathList(node, path)).toBeUndefined();
  });

  it('returns the original value for an empty path', () => {
    const node = { value: 'unchanged' };
    expect(getTextByPathList(node, [])).toBe(node);
  });

  it('normalizes a nullish terminal value to undefined', () => {
    expect(getTextByPathList({ root: null }, ['root'])).toBeUndefined();
    expect(getTextByPathList({ root: undefined }, ['root'])).toBeUndefined();
  });

  it('iterates a singleton with index zero', () => {
    const callback = vi.fn(
      (item: unknown, index: number) => `${index}:${String(item)}`,
    );
    expect(eachElement('one', callback)).toBe('0:one');
    expect(callback).toHaveBeenCalledOnce();
  });

  it('iterates arrays in order and concatenates callback results', () => {
    const calls: Array<[unknown, number]> = [];
    const output = eachElement(['a', 'b', 'c'], (item, index) => {
      calls.push([item, index]);
      return index === 1 ? false : `${index}:${String(item)}`;
    });

    expect(output).toBe('0:afalse2:c');
    expect(calls).toEqual([
      ['a', 0],
      ['b', 1],
      ['c', 2],
    ]);
  });

  it('preserves nullish input and defines empty-array output', () => {
    const callback = vi.fn(() => 'unused');
    expect(eachElement(null, callback)).toBeNull();
    expect(eachElement(undefined, callback)).toBeUndefined();
    expect(eachElement([], callback)).toBe('');
    expect(callback).not.toHaveBeenCalled();
  });

  it('stringifies nullish callback results deterministically', () => {
    expect(
      eachElement([1, 2], (_item, index) => (index ? null : undefined)),
    ).toBe('undefinednull');
  });
});
