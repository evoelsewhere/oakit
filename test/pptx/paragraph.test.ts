import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import {
  getHorizontalAlign,
  getParagraphIndent,
  getParagraphSpacing,
  getTextAutoFit,
  getVerticalAlign,
} from '../../src/formats/pptx/internal/paragraph';

interface ParagraphFixture {
  defaultTextStyle?: XmlLookupValue;
  layout?: XmlLookupValue;
  master?: XmlLookupValue;
  masterTextStyles?: XmlLookupValue;
  paragraph?: XmlLookupValue;
  shape?: XmlLookupValue;
  textBody?: XmlLookupValue;
  type?: string;
}

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(fixture: ParagraphFixture = {}): PptxParserContext {
  return {
    defaultTextStyle: fixture.defaultTextStyle ?? xml({}),
    slideMasterTextStyles: fixture.masterTextStyles ?? xml({}),
  } as unknown as PptxParserContext;
}

function paragraphStyle(
  attrs: Record<string, string> = {},
  children: Record<string, object> = {},
): XmlLookupValue {
  return xml({ attrs, ...children });
}

function paragraph(style: XmlLookupValue): XmlLookupValue {
  return xml({ 'a:pPr': style });
}

function levelStyle(style: XmlLookupValue, level = 1): XmlLookupValue {
  return xml({ 'a:lstStyle': { [`a:lvl${level}pPr`]: style } });
}

function shapeLevelStyle(style: XmlLookupValue, level = 1): XmlLookupValue {
  return xml({ 'p:txBody': levelStyle(style, level) });
}

function shapeParagraphStyle(style: XmlLookupValue): XmlLookupValue {
  return xml({ 'p:txBody': { 'a:p': { 'a:pPr': style } } });
}

function masterTextStyle(
  section: 'p:bodyStyle' | 'p:otherStyle' | 'p:titleStyle',
  style: XmlLookupValue,
  level = 1,
): XmlLookupValue {
  return xml({ [section]: { [`a:lvl${level}pPr`]: style } });
}

function horizontal(fixture: ParagraphFixture = {}): string {
  return getHorizontalAlign(
    fixture.paragraph ?? xml({}),
    fixture.shape ?? xml({}),
    fixture.type ?? 'body',
    fixture.layout,
    fixture.master,
    context(fixture),
  );
}

function spacing(fixture: ParagraphFixture = {}) {
  return getParagraphSpacing(
    fixture.paragraph ?? xml({}),
    fixture.textBody ?? xml({}),
    fixture.layout,
    fixture.master,
    fixture.type ?? 'body',
    fixture.masterTextStyles,
    context(fixture),
  );
}

function indent(fixture: ParagraphFixture = {}) {
  return getParagraphIndent(
    fixture.paragraph ?? xml({}),
    fixture.textBody ?? xml({}),
    fixture.layout,
    fixture.master,
    fixture.type ?? 'body',
    fixture.masterTextStyles,
    context(fixture),
  );
}

function spacingChild(
  kind: 'a:lnSpc' | 'a:spcAft' | 'a:spcBef',
  unit: 'a:spcPct' | 'a:spcPts',
  value: string,
): Record<string, object> {
  return { [kind]: { [unit]: { attrs: { val: value } } } };
}

describe('PowerPoint paragraph alignment', () => {
  it.each([
    ['l', 'left'],
    ['r', 'right'],
    ['ctr', 'center'],
    ['just', 'justify'],
    ['dist', 'justify'],
    ['justLow', 'justify'],
    ['thaiDist', 'justify'],
    ['invalid', 'inherit'],
  ])('maps DrawingML alignment %s to %s', (value, expected) => {
    expect(
      horizontal({ paragraph: paragraph(paragraphStyle({ algn: value })) }),
    ).toBe(expected);
  });

  it('defaults an absent alignment to left', () => {
    expect(horizontal()).toBe('left');
  });

  it.each([
    [
      'shape paragraph',
      { shape: shapeParagraphStyle(paragraphStyle({ algn: 'r' })) },
      'right',
    ],
    [
      'layout list level',
      { layout: shapeLevelStyle(paragraphStyle({ algn: 'ctr' })) },
      'center',
    ],
    [
      'layout paragraph',
      { layout: shapeParagraphStyle(paragraphStyle({ algn: 'r' })) },
      'right',
    ],
    [
      'master list level',
      { master: shapeLevelStyle(paragraphStyle({ algn: 'ctr' })) },
      'center',
    ],
    [
      'master paragraph',
      { master: shapeParagraphStyle(paragraphStyle({ algn: 'r' })) },
      'right',
    ],
    [
      'master title style',
      {
        masterTextStyles: masterTextStyle(
          'p:titleStyle',
          paragraphStyle({ algn: 'ctr' }),
        ),
        type: 'title',
      },
      'center',
    ],
    [
      'master centered-title style',
      {
        masterTextStyles: masterTextStyle(
          'p:titleStyle',
          paragraphStyle({ algn: 'ctr' }),
        ),
        type: 'ctrTitle',
      },
      'center',
    ],
    [
      'master subtitle body fallback',
      {
        masterTextStyles: masterTextStyle(
          'p:bodyStyle',
          paragraphStyle({ algn: 'r' }),
        ),
        type: 'subTitle',
      },
      'right',
    ],
    [
      'master body style',
      {
        masterTextStyles: masterTextStyle(
          'p:bodyStyle',
          paragraphStyle({ algn: 'just' }),
        ),
      },
      'justify',
    ],
    [
      'master other style',
      {
        masterTextStyles: masterTextStyle(
          'p:otherStyle',
          paragraphStyle({ algn: 'r' }),
        ),
        type: 'obj',
      },
      'right',
    ],
  ] as const)('inherits alignment from %s', (_name, fixture, expected) => {
    expect(horizontal(fixture)).toBe(expected);
  });

  it('uses the first authored alignment in the inheritance chain', () => {
    const fixture: ParagraphFixture = {
      paragraph: paragraph(paragraphStyle({ algn: 'l' })),
      shape: shapeParagraphStyle(paragraphStyle({ algn: 'r' })),
      layout: shapeLevelStyle(paragraphStyle({ algn: 'ctr' })),
      master: shapeLevelStyle(paragraphStyle({ algn: 'just' })),
      masterTextStyles: masterTextStyle(
        'p:bodyStyle',
        paragraphStyle({ algn: 'dist' }),
      ),
    };

    expect(horizontal(fixture)).toBe('left');
  });

  it('inherits past a layout node without an authored alignment', () => {
    expect(
      horizontal({
        layout: xml({ 'p:txBody': { 'a:bodyPr': {} } }),
        master: shapeLevelStyle(paragraphStyle({ algn: 'r' })),
      }),
    ).toBe('right');
  });

  it.each([
    ['0', 1, 'center'],
    ['1', 2, 'center'],
    ['2', 3, 'center'],
    ['3', 4, 'center'],
    ['4', 5, 'center'],
    ['5', 6, 'center'],
    ['6', 7, 'center'],
    ['7', 8, 'center'],
    ['8', 9, 'center'],
    ['0junk', 1, 'center'],
    ['-1', 1, 'center'],
    ['9', 1, 'center'],
    ['1.5', 1, 'center'],
  ])(
    'normalizes paragraph level %j to level path %s',
    (value, level, expected) => {
      expect(
        horizontal({
          paragraph: paragraph(paragraphStyle({ lvl: value })),
          layout: shapeLevelStyle(paragraphStyle({ algn: 'ctr' }), level),
        }),
      ).toBe(expected);
    },
  );

  it.each(['title', 'ctrTitle'])(
    'does not inherit a body alignment for %s placeholders',
    (type) => {
      expect(
        horizontal({
          masterTextStyles: masterTextStyle(
            'p:bodyStyle',
            paragraphStyle({ algn: 'r' }),
          ),
          type,
        }),
      ).toBe('left');
    },
  );
});

describe('PowerPoint vertical alignment and autofit', () => {
  it.each([
    ['t', 'up'],
    ['ctr', 'mid'],
    ['b', 'down'],
    ['just', 'up'],
    ['dist', 'up'],
    ['invalid', 'up'],
  ])('maps vertical anchor %s to %s', (anchor, expected) => {
    const node = xml({
      'p:txBody': { 'a:bodyPr': { attrs: { anchor } } },
    });

    expect(getVerticalAlign(node, undefined, undefined)).toBe(expected);
  });

  it('inherits vertical alignment from layout then master and defaults to top', () => {
    const layout = xml({
      'p:txBody': { 'a:bodyPr': { attrs: { anchor: 'ctr' } } },
    });
    const master = xml({
      'p:txBody': { 'a:bodyPr': { attrs: { anchor: 'b' } } },
    });

    expect(getVerticalAlign(xml({}), layout, master)).toBe('mid');
    expect(getVerticalAlign(xml({}), undefined, master)).toBe('down');
    expect(getVerticalAlign(xml({}), undefined, undefined)).toBe('up');
  });

  it('lets explicit no-autofit stop inherited autofit', () => {
    const node = xml({
      'p:txBody': { 'a:bodyPr': { 'a:noAutofit': {} } },
    });
    const layout = xml({
      'p:txBody': { 'a:bodyPr': { 'a:spAutoFit': {} } },
    });

    expect(getTextAutoFit(node, layout, undefined)).toBeNull();
  });

  it('resolves shape autofit from slide, layout, and master', () => {
    const fit = xml({
      'p:txBody': { 'a:bodyPr': { 'a:spAutoFit': {} } },
    });

    expect(getTextAutoFit(fit, undefined, undefined)).toEqual({
      type: 'shape',
    });
    expect(getTextAutoFit(xml({}), fit, undefined)).toEqual({ type: 'shape' });
    expect(getTextAutoFit(xml({}), undefined, fit)).toEqual({ type: 'shape' });
  });

  it('inherits autofit past an empty local body property node', () => {
    const node = xml({ 'p:txBody': { 'a:bodyPr': {} } });
    const layout = xml({
      'p:txBody': { 'a:bodyPr': { 'a:spAutoFit': {} } },
    });

    expect(getTextAutoFit(node, layout, undefined)).toEqual({ type: 'shape' });
  });

  it.each([
    ['1000', { type: 'text', fontScale: 1 }],
    ['75000', { type: 'text', fontScale: 75 }],
    ['100000', { type: 'text', fontScale: 100 }],
    ['', { type: 'text' }],
    ['999', { type: 'text' }],
    ['100001', { type: 'text' }],
    ['75000junk', { type: 'text' }],
    ['junk75000', { type: 'text' }],
    ['1e4', { type: 'text' }],
  ] as const)('normalizes text autofit font scale %j', (value, expected) => {
    const node = xml({
      'p:txBody': {
        'a:bodyPr': {
          'a:normAutofit': value ? { attrs: { fontScale: value } } : {},
        },
      },
    });

    expect(getTextAutoFit(node, undefined, undefined)).toEqual(expected);
  });

  it('uses no-autofit before shape and text autofit in one malformed body', () => {
    const node = xml({
      'p:txBody': {
        'a:bodyPr': {
          'a:noAutofit': {},
          'a:spAutoFit': {},
          'a:normAutofit': { attrs: { fontScale: '75000' } },
        },
      },
    });

    expect(getTextAutoFit(node, undefined, undefined)).toBeNull();
    expect(getTextAutoFit(xml({}), undefined, undefined)).toBeNull();
  });
});

describe('PowerPoint paragraph spacing and indentation', () => {
  it('converts percentage and point spacing exactly', () => {
    const style = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPct', '150000'),
        ...spacingChild('a:spcBef', 'a:spcPct', '50000'),
        ...spacingChild('a:spcAft', 'a:spcPts', '1200'),
      },
    );

    expect(spacing({ paragraph: paragraph(style) })).toEqual({
      lineSpacing: 1.5,
      spaceBefore: '0.5em',
      spaceAfter: '12pt',
    });
  });

  it('converts point line spacing and percentage after-spacing', () => {
    const style = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPts', '1800'),
        ...spacingChild('a:spcAft', 'a:spcPct', '25000'),
      },
    );

    expect(spacing({ paragraph: paragraph(style) })).toEqual({
      lineSpacing: '18pt',
      spaceAfter: '0.25em',
    });
  });

  it('preserves explicit zero spacing', () => {
    const style = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPct', '0'),
        ...spacingChild('a:spcBef', 'a:spcPts', '0'),
        ...spacingChild('a:spcAft', 'a:spcPct', '0'),
      },
    );

    expect(spacing({ paragraph: paragraph(style) })).toEqual({
      lineSpacing: 0,
      spaceBefore: '0pt',
      spaceAfter: '0em',
    });
  });

  it('preserves explicit zero point line spacing and rejects negative points', () => {
    const zero = paragraphStyle({}, spacingChild('a:lnSpc', 'a:spcPts', '0'));
    const negative = paragraphStyle(
      {},
      spacingChild('a:lnSpc', 'a:spcPts', '-1'),
    );

    expect(spacing({ paragraph: paragraph(zero) })).toEqual({
      lineSpacing: '0pt',
    });
    expect(spacing({ paragraph: paragraph(negative) })).toBeNull();
  });

  it.each([
    '100000junk',
    'junk100000',
    '1e5',
    'Infinity',
    '-1',
    '9007199254740992',
  ])('rejects malformed or invalid spacing %j', (value) => {
    const style = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPct', value),
        ...spacingChild('a:spcBef', 'a:spcPts', value),
        ...spacingChild('a:spcAft', 'a:spcPct', value),
      },
    );

    expect(spacing({ paragraph: paragraph(style) })).toBeNull();
  });

  it('resolves each spacing field independently across inheritance layers', () => {
    const fixture: ParagraphFixture = {
      paragraph: paragraph(
        paragraphStyle({}, spacingChild('a:lnSpc', 'a:spcPct', '120000')),
      ),
      textBody: levelStyle(
        paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '600')),
      ),
      layout: shapeLevelStyle(
        paragraphStyle({}, spacingChild('a:spcAft', 'a:spcPct', '25000')),
      ),
      master: shapeLevelStyle(
        paragraphStyle({}, spacingChild('a:spcAft', 'a:spcPts', '9900')),
      ),
    };

    expect(spacing(fixture)).toEqual({
      lineSpacing: 1.2,
      spaceBefore: '6pt',
      spaceAfter: '0.25em',
    });
  });

  it('does not let lower-priority styles overwrite resolved spacing', () => {
    const high = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPct', '120000'),
        ...spacingChild('a:spcBef', 'a:spcPts', '600'),
        ...spacingChild('a:spcAft', 'a:spcPts', '700'),
      },
    );
    const low = paragraphStyle(
      {},
      {
        ...spacingChild('a:lnSpc', 'a:spcPct', '200000'),
        ...spacingChild('a:spcBef', 'a:spcPts', '1600'),
        ...spacingChild('a:spcAft', 'a:spcPts', '1700'),
      },
    );

    expect(
      spacing({
        paragraph: paragraph(high),
        textBody: levelStyle(low),
      }),
    ).toEqual({
      lineSpacing: 1.2,
      spaceBefore: '6pt',
      spaceAfter: '7pt',
    });
  });

  it.each([
    [
      'layout paragraph',
      {
        layout: shapeParagraphStyle(
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '700')),
        ),
      },
      '7pt',
    ],
    [
      'master paragraph',
      {
        master: shapeParagraphStyle(
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '800')),
        ),
      },
      '8pt',
    ],
    [
      'master title text style',
      {
        masterTextStyles: masterTextStyle(
          'p:titleStyle',
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '900')),
        ),
        type: 'title',
      },
      '9pt',
    ],
    [
      'master centered-title text style',
      {
        masterTextStyles: masterTextStyle(
          'p:titleStyle',
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '950')),
        ),
        type: 'ctrTitle',
      },
      '9.5pt',
    ],
    [
      'master subtitle body fallback',
      {
        masterTextStyles: masterTextStyle(
          'p:bodyStyle',
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '1000')),
        ),
        type: 'subTitle',
      },
      '10pt',
    ],
    [
      'master body text style',
      {
        masterTextStyles: masterTextStyle(
          'p:bodyStyle',
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '1100')),
        ),
      },
      '11pt',
    ],
    [
      'master other text style',
      {
        masterTextStyles: masterTextStyle(
          'p:otherStyle',
          paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '1200')),
        ),
        type: 'obj',
      },
      '12pt',
    ],
    [
      'level default text style',
      {
        defaultTextStyle: xml({
          'a:lvl1pPr': paragraphStyle(
            {},
            spacingChild('a:spcBef', 'a:spcPts', '1300'),
          ),
        }),
      },
      '13pt',
    ],
    [
      'default paragraph style',
      {
        defaultTextStyle: xml({
          'a:defPPr': paragraphStyle(
            {},
            spacingChild('a:spcBef', 'a:spcPts', '1400'),
          ),
        }),
      },
      '14pt',
    ],
  ] as const)('inherits spacing from %s', (_name, fixture, expected) => {
    expect(spacing(fixture)?.spaceBefore).toBe(expected);
  });

  it.each(['title', 'ctrTitle'])(
    'does not use body paragraph spacing for %s placeholders',
    (type) => {
      expect(
        spacing({
          masterTextStyles: masterTextStyle(
            'p:bodyStyle',
            paragraphStyle({}, spacingChild('a:spcBef', 'a:spcPts', '900')),
          ),
          type,
        }),
      ).toBeNull();
    },
  );

  it('converts authored indentation and preserves zero', () => {
    expect(
      indent({
        paragraph: paragraph(
          paragraphStyle({ indent: '-6350', marL: '12700' }),
        ),
      }),
    ).toEqual({ marginLeft: '1pt', textIndent: '-0.5pt' });
    expect(
      indent({
        paragraph: paragraph(paragraphStyle({ indent: '0', marL: '0' })),
      }),
    ).toEqual({ marginLeft: '0pt', textIndent: '0pt' });
  });

  it.each(['12700junk', 'junk12700', '1e4', 'Infinity', '9007199254740992'])(
    'rejects malformed indentation %j',
    (value) => {
      expect(
        indent({
          paragraph: paragraph(paragraphStyle({ indent: value, marL: value })),
        }),
      ).toBeNull();
    },
  );

  it('resolves indentation fields independently across inheritance layers', () => {
    expect(
      indent({
        paragraph: paragraph(paragraphStyle({ marL: '12700' })),
        textBody: levelStyle(paragraphStyle({ indent: '-6350' })),
        layout: shapeLevelStyle(
          paragraphStyle({ indent: '12700', marL: '25400' }),
        ),
      }),
    ).toEqual({ marginLeft: '1pt', textIndent: '-0.5pt' });
  });

  it('returns null when no spacing or indentation is authored', () => {
    expect(spacing()).toBeNull();
    expect(indent()).toBeNull();
  });
});
