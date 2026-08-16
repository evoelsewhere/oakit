import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';

import {
  serializeMinimalSlideLayout,
  serializeMinimalSlideMaster,
} from '../../src/formats/pptx/writer/hierarchy';

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

function elementByName(
  values: readonly ParsedElement[],
  name: string,
): ParsedElement {
  const value = values.find((element) => element.name === name);
  if (!value) throw new Error(`Missing test element: ${name}`);
  return value;
}

async function sha256(value: string): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest(
    'SHA-256',
    new TextEncoder().encode(value),
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, '0'),
  ).join('');
}

describe('PowerPoint minimal hierarchy serialization', () => {
  it('serializes a complete master color mapping', () => {
    const parsed = elements(serializeMinimalSlideMaster());
    const colorMap = elementByName(parsed, 'p:clrMap');

    expect(parsed[0]?.name).toBe('p:sldMaster');
    expect(colorMap.attributes).toEqual({
      accent1: 'accent1',
      accent2: 'accent2',
      accent3: 'accent3',
      accent4: 'accent4',
      accent5: 'accent5',
      accent6: 'accent6',
      bg1: 'lt1',
      bg2: 'lt2',
      folHlink: 'folHlink',
      hlink: 'hlink',
      tx1: 'dk1',
      tx2: 'dk2',
    });
  });

  it('binds the master to one canonical layout relationship', () => {
    const layoutId = elementByName(
      elements(serializeMinimalSlideMaster()),
      'p:sldLayoutId',
    );

    expect(layoutId.attributes.id).toBe('2147483649');
    expect(layoutId.attributes['r:id']).toBe('rId1');
  });

  it('serializes a preserved blank layout with master color mapping', () => {
    const parsed = elements(serializeMinimalSlideLayout());
    const root = parsed[0];

    expect(root?.name).toBe('p:sldLayout');
    expect(root?.attributes.type).toBe('blank');
    expect(root?.attributes.preserve).toBe('1');
    expect(parsed.map((element) => element.name)).toContain(
      'a:masterClrMapping',
    );
  });

  it.each([
    ['master', serializeMinimalSlideMaster()],
    ['layout', serializeMinimalSlideLayout()],
  ])('uses a complete empty shape-tree root for the %s', (_owner, xml) => {
    const names = elements(xml).map((element) => element.name);

    expect(names).toEqual(
      expect.arrayContaining([
        'p:spTree',
        'p:nvGrpSpPr',
        'p:cNvPr',
        'p:cNvGrpSpPr',
        'p:nvPr',
        'p:grpSpPr',
        'a:xfrm',
        'a:off',
        'a:ext',
        'a:chOff',
        'a:chExt',
      ]),
    );
    expect(elementByName(elements(xml), 'p:cNvPr').attributes).toEqual({
      id: '1',
      name: '',
    });
  });

  it('locks the reviewed hierarchy bytes', async () => {
    expect(await sha256(serializeMinimalSlideMaster())).toBe(
      '391ed3ffec61406d7c5157f7cbd1875cc84bc2abbbc0c5e0c5af34858b675c3c',
    );
    expect(await sha256(serializeMinimalSlideLayout())).toBe(
      '91cab059103522716b55ca2385db488dba70d9a9dce5c772aff4bdc30b3602ca',
    );
  });
});
