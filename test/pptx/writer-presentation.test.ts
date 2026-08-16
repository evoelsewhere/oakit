import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';

import { serializePresentation } from '../../src/formats/pptx/writer/presentation';

interface ParsedElement {
  attributes: Record<string, string>;
  name: string;
}

function elements(xml: string): ParsedElement[] {
  const result: ParsedElement[] = [];
  const parser = new SaxesParser({ xmlns: true });
  parser.on('opentag', (tag) => {
    const attributes: Record<string, string> = {};
    for (const [name, attribute] of Object.entries(tag.attributes)) {
      attributes[name] = attribute.value;
    }
    result.push({ attributes, name: tag.name });
  });
  parser.write(xml).close();
  return result;
}

function byName(values: readonly ParsedElement[], name: string) {
  return values.filter((element) => element.name === name);
}

describe('PowerPoint presentation serialization', () => {
  it('serializes a zero-slide presentation without an empty slide list', () => {
    const xml = serializePresentation({ height: 540, width: 960 }, 0);
    const parsed = elements(xml);

    expect(parsed[0]?.name).toBe('p:presentation');
    expect(byName(parsed, 'p:sldIdLst')).toEqual([]);
    expect(byName(parsed, 'p:sldId')).toEqual([]);
    expect(byName(parsed, 'p:sldMasterId')).toEqual([
      {
        attributes: { id: '2147483648', 'r:id': 'rId1' },
        name: 'p:sldMasterId',
      },
    ]);
    expect(xml).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="12192000" cy="6858000" type="custom"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle/></p:presentation>',
    );
  });

  it('allocates ordered slide and relationship identities', () => {
    const parsed = elements(
      serializePresentation({ height: 540, width: 960 }, 3),
    );

    expect(byName(parsed, 'p:sldId').map((value) => value.attributes)).toEqual([
      { id: '256', 'r:id': 'rId2' },
      { id: '257', 'r:id': 'rId3' },
      { id: '258', 'r:id': 'rId4' },
    ]);
    expect(serializePresentation({ height: 540, width: 960 }, 3)).toContain(
      'r:id="rId2"/><p:sldId id="257"',
    );
  });

  it('converts slide dimensions to exact EMUs and declares custom size', () => {
    const parsed = elements(
      serializePresentation({ height: 540.5, width: 960.25 }, 1),
    );

    expect(byName(parsed, 'p:sldSz')[0]?.attributes).toEqual({
      cx: '12195175',
      cy: '6864350',
      type: 'custom',
    });
  });

  it('uses a stable portrait notes-page size and default text style', () => {
    const parsed = elements(
      serializePresentation({ height: 540, width: 960 }, 1),
    );

    expect(byName(parsed, 'p:notesSz')[0]?.attributes).toEqual({
      cx: '6858000',
      cy: '9144000',
    });
    expect(byName(parsed, 'p:defaultTextStyle')).toHaveLength(1);
  });

  it.each([-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 10_001])(
    'rejects invalid presentation slide count %s',
    (slideCount) => {
      expect(() =>
        serializePresentation({ height: 540, width: 960 }, slideCount),
      ).toThrow(
        new RangeError(
          'PowerPoint presentation slide count must be an integer from 0 through 10000',
        ),
      );
    },
  );

  it('accepts the bounded maximum without overflowing slide identities', () => {
    const xml = serializePresentation({ height: 540, width: 960 }, 10_000);

    expect(xml).toContain('<p:sldId id="10255" r:id="rId10001"/>');
    expect(xml.startsWith('<?xml version="1.0"')).toBe(true);
  });
});
