import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';

import { serializeMinimalTheme } from '../../src/formats/pptx/writer/theme';

function elementNames(xml: string): string[] {
  const names: string[] = [];
  const parser = new SaxesParser({ xmlns: true });
  parser.on('opentag', (tag) => names.push(tag.name));
  parser.write(xml).close();
  return names;
}

function childrenBetween(
  names: readonly string[],
  start: string,
  end: string,
  child: string,
): number {
  const startIndex = names.indexOf(start);
  const endIndex = names.indexOf(end, startIndex + 1);
  return names.slice(startIndex + 1, endIndex).filter((name) => name === child)
    .length;
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

describe('PowerPoint minimal theme serialization', () => {
  it('is a standalone DrawingML theme with no external references', () => {
    const xml = serializeMinimalTheme();

    expect(xml).toContain(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="OAKit Minimal">',
    );
    expect(xml.endsWith('</a:theme>')).toBe(true);
    expect(xml).not.toContain('r:id=');
    expect(xml).not.toContain('Target=');
    expect(() => elementNames(xml)).not.toThrow();
  });

  it('contains every required color-scheme slot in schema order', () => {
    const names = elementNames(serializeMinimalTheme());
    const colorNames = [
      'a:dk1',
      'a:lt1',
      'a:dk2',
      'a:lt2',
      'a:accent1',
      'a:accent2',
      'a:accent3',
      'a:accent4',
      'a:accent5',
      'a:accent6',
      'a:hlink',
      'a:folHlink',
    ];

    expect(names.filter((name) => colorNames.includes(name))).toEqual(
      colorNames,
    );
  });

  it('contains complete major and minor font collections', () => {
    const xml = serializeMinimalTheme();

    expect(xml).toContain(
      '<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>',
    );
    expect(xml).toContain(
      '<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>',
    );
  });

  it('contains the required three entries in every style-matrix list', () => {
    const names = elementNames(serializeMinimalTheme());

    expect(
      childrenBetween(names, 'a:fillStyleLst', 'a:lnStyleLst', 'a:solidFill'),
    ).toBe(3);
    expect(
      childrenBetween(names, 'a:lnStyleLst', 'a:effectStyleLst', 'a:ln'),
    ).toBe(3);
    expect(
      childrenBetween(
        names,
        'a:effectStyleLst',
        'a:bgFillStyleLst',
        'a:effectStyle',
      ),
    ).toBe(3);
    expect(
      names
        .slice(names.indexOf('a:bgFillStyleLst') + 1)
        .filter((name) => name === 'a:solidFill'),
    ).toHaveLength(3);
  });

  it('uses distinct canonical line widths', () => {
    const xml = serializeMinimalTheme();

    expect(xml.match(/<a:ln w="6350"/g)).toHaveLength(1);
    expect(xml.match(/<a:ln w="12700"/g)).toHaveLength(1);
    expect(xml.match(/<a:ln w="19050"/g)).toHaveLength(1);
  });

  it('locks the reviewed canonical theme bytes', async () => {
    expect(await sha256(serializeMinimalTheme())).toBe(
      'e40ca1b4afc65ccb3a3d7b50f69fdc119572a5a4a9aecf86cb3aa345e0d9f66f',
    );
  });
});
