import { describe, expect, it } from 'vitest';

import { serializeRelationships } from '../../src/formats/pptx/writer/relationships';

const PREFIX =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

describe('PowerPoint relationship serialization', () => {
  it('serializes an empty relationship part exactly', () => {
    expect(serializeRelationships([])).toBe(`${PREFIX}</Relationships>`);
  });

  it('preserves owner-scoped relationship order', () => {
    expect(
      serializeRelationships([
        {
          id: 'rId1',
          target: 'slides/slide1.xml',
          type: 'http://example.test/slide',
        },
        {
          id: 'rId2',
          target: 'slideMasters/slideMaster1.xml',
          type: 'http://example.test/master',
        },
      ]),
    ).toBe(
      `${PREFIX}<Relationship Id="rId1" Type="http://example.test/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://example.test/master" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
    );
  });

  it('escapes every caller-provided relationship attribute', () => {
    expect(
      serializeRelationships([
        {
          id: `id<&"' _x0041_`,
          target: `target<&"' _x0042_`,
          type: `type<&"' _x0043_`,
        },
      ]),
    ).toBe(
      `${PREFIX}<Relationship Id="id&lt;&amp;&quot;&apos; _x005F_x0041_" Type="type&lt;&amp;&quot;&apos; _x005F_x0043_" Target="target&lt;&amp;&quot;&apos; _x005F_x0042_"/></Relationships>`,
    );
  });

  it('emits external mode only when explicitly requested', () => {
    const internal = serializeRelationships([
      { id: 'rId1', target: 'internal.xml', type: 'internal' },
    ]);
    const external = serializeRelationships([
      {
        id: 'rId1',
        target: 'https://example.test',
        targetMode: 'External',
        type: 'external',
      },
    ]);

    expect(internal).not.toContain('TargetMode=');
    expect(external).toContain(' TargetMode="External"/>');
  });
});
