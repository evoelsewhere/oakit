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

  it('preserves the inverted-line preset as stroke-only geometry', async () => {
    const input = await createMinimalPptx({
      'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8"?>
        <p:sld>
          <p:cSld>
            <p:spTree>
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="2" name="Inverted line"/>
                  <p:cNvSpPr/>
                  <p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="914400" y="914400"/>
                    <a:ext cx="1828800" cy="914400"/>
                  </a:xfrm>
                  <a:prstGeom prst="lineInv"><a:avLst/></a:prstGeom>
                  <a:noFill/>
                  <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                </p:spPr>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sld>`,
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.slides[0]?.elements[0]).toMatchObject({
      type: 'shape',
      shapType: 'lineInv',
      path: 'M 0 72 L 144 0',
      strokeOnly: true,
    });
  });
});
