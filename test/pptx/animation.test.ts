import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import {
  findTransitionNode,
  parseTransition,
} from '../../src/formats/pptx/internal/animation';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

describe('PPTX transition parsing', () => {
  it('returns null for missing content, root names, and transition nodes', () => {
    expect(findTransitionNode(null, 'p:sld')).toBeNull();
    expect(findTransitionNode(xml({}), '')).toBeNull();
    expect(findTransitionNode(xml({ 'p:sld': {} }), 'p:sld')).toBeNull();
    expect(parseTransition(null)).toBeNull();
  });

  it('prefers a direct transition over alternate content', () => {
    const direct = xml({ attrs: { spd: 'fast' } });
    const choice = xml({ attrs: { spd: 'med' } });
    const fallback = xml({ attrs: { spd: 'slow' } });
    const content = xml({
      'p:sld': {
        'mc:AlternateContent': {
          'mc:Choice': { 'p:transition': choice },
          'mc:Fallback': { 'p:transition': fallback },
        },
        'p:transition': direct,
      },
    });

    expect(findTransitionNode(content, 'p:sld')).toBe(direct);
  });

  it('uses alternate choice before fallback and supports fallback alone', () => {
    const choice = xml({ attrs: { spd: 'med' } });
    const fallback = xml({ attrs: { spd: 'slow' } });

    expect(
      findTransitionNode(
        xml({
          'p:sldLayout': {
            'mc:AlternateContent': {
              'mc:Choice': { 'p:transition': choice },
              'mc:Fallback': { 'p:transition': fallback },
            },
          },
        }),
        'p:sldLayout',
      ),
    ).toBe(choice);
    expect(
      findTransitionNode(
        xml({
          'p:sldMaster': {
            'mc:AlternateContent': {
              'mc:Fallback': { 'p:transition': fallback },
            },
          },
        }),
        'p:sldMaster',
      ),
    ).toBe(fallback);
  });

  it('defines stable defaults for an empty transition', () => {
    expect(parseTransition(xml({}))).toEqual({
      direction: null,
      duration: 1000,
      type: 'none',
    });
  });

  it.each([
    ['slow', 1000],
    ['med', 800],
    ['fast', 500],
    ['unknown', 1000],
  ])('maps transition speed %s to %s ms', (speed, duration) => {
    expect(parseTransition(xml({ attrs: { spd: speed } }))).toMatchObject({
      duration,
    });
  });

  it('prefers a valid extension duration over speed and effect duration', () => {
    expect(
      parseTransition(
        xml({
          attrs: { 'p14:dur': '1750', spd: 'fast' },
          'p:wipe': { attrs: { dur: '2500', dir: 'l' } },
        }),
      ),
    ).toEqual({ direction: 'l', duration: 1750, type: 'wipe' });
  });

  it('requires the complete extension duration attribute name', () => {
    expect(
      parseTransition(
        xml({
          attrs: {
            'p14:duration': '3000',
            spd: 'fast',
            'xp14:dur': '2000',
          },
        }),
      ),
    ).toMatchObject({ duration: 500 });
  });

  it('uses a valid effect duration when no extension duration exists', () => {
    expect(
      parseTransition(
        xml({
          attrs: { spd: 'fast' },
          'p14:morph': { attrs: { dur: '2250', dir: 'in' } },
        }),
      ),
    ).toEqual({ direction: 'in', duration: 2250, type: 'morph' });
  });

  it.each(['1750x', '-1', '1.5', 'Infinity', ''])(
    'rejects malformed extension duration %j without partial parsing',
    (duration) => {
      expect(
        parseTransition(xml({ attrs: { 'p14:dur': duration, spd: 'fast' } })),
      ).toMatchObject({ duration: 500 });
    },
  );

  it.each(['2250x', '-1', '1.5', 'Infinity', ''])(
    'rejects malformed effect duration %j without partial parsing',
    (duration) => {
      expect(
        parseTransition(xml({ 'p:wipe': { attrs: { dur: duration } } })),
      ).toEqual({ direction: null, duration: 1000, type: 'wipe' });
    },
  );

  it('sets automatic advance only when click advance is disabled', () => {
    expect(
      parseTransition(xml({ attrs: { advClick: '0', advTm: '2500' } })),
    ).toMatchObject({ autoNextAfter: 2500 });
    expect(
      parseTransition(xml({ attrs: { advClick: '1', advTm: '2500' } })),
    ).not.toHaveProperty('autoNextAfter');
  });

  it.each(['2500x', '-1', '1.5', 'Infinity', ''])(
    'omits malformed automatic advance %j',
    (advance) => {
      expect(
        parseTransition(xml({ attrs: { advClick: '0', advTm: advance } })),
      ).not.toHaveProperty('autoNextAfter');
    },
  );

  it('ignores non-transition children and selects the first effect', () => {
    expect(
      parseTransition(
        xml({
          'a:ext': { attrs: { dir: 'ignored' } },
          attrs: {},
          metadata: 'ignored',
          'xp:cover': { attrs: { dir: 'ignored' } },
          'p:push': { attrs: { dir: 'r' } },
          'p:wipe': { attrs: { dir: 'l' } },
        }),
      ),
    ).toEqual({ direction: 'r', duration: 1000, type: 'push' });
  });
});
