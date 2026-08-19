import { describe, expect, it } from 'vitest';

import type {
  PptxSceneGroupElement,
  PptxSceneGroupTransform,
} from '../../src/formats/pptx/scene-types';
import {
  serializeGroup,
  serializeGroupTransform,
} from '../../src/formats/pptx/writer/group';

const TRANSFORM: PptxSceneGroupTransform = {
  childSpace: { height: 100, width: 200, x: 5, y: 10 },
  height: 150,
  width: 300,
  x: 20,
  y: 30,
};

function group(): PptxSceneGroupElement {
  return {
    authored: {},
    elements: [],
    key: 'group-1',
    resolved: { hidden: false },
    type: 'group',
  };
}

describe('native PowerPoint group serialization', () => {
  it('serializes outer and child coordinate spaces in exact EMUs', () => {
    expect(serializeGroupTransform(TRANSFORM)).toBe(
      '<a:xfrm><a:off x="254000" y="381000"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="63500" y="127000"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm>',
    );
  });

  it('serializes rotation, explicit flips, metadata, and owned children', () => {
    const element = group();
    element.name = `Group <&"'`;
    element.description = `Description <&"'`;
    element.title = `Title <&"'`;
    element.authored.hidden = false;
    const xml = serializeGroup(
      element,
      {
        ...TRANSFORM,
        flipHorizontal: false,
        flipVertical: true,
        rotation: -15,
      },
      7,
      '<p:sp id="child"/>',
    );

    expect(xml).toContain(
      '<p:cNvPr id="7" name="Group &lt;&amp;&quot;&apos;" descr="Description &lt;&amp;&quot;&apos;" title="Title &lt;&amp;&quot;&apos;" hidden="0"/>',
    );
    expect(xml).toContain('<a:xfrm rot="-900000" flipH="0" flipV="1">');
    expect(xml).toContain('</p:grpSpPr><p:sp id="child"/></p:grpSp>');
  });

  it('uses a deterministic default group name', () => {
    expect(serializeGroup(group(), TRANSFORM, 4, '')).toContain(
      '<p:cNvPr id="4" name="Group 4"/>',
    );
  });
});
