import { describe, expect, it, vi } from 'vitest';

import type { PptxSceneTextBody } from '../../src/formats/pptx/scene-types';
import {
  serializeTextBody,
  serializeTextBodyProperties,
} from '../../src/formats/pptx/writer/text-body';
import type { PptxTextSerializationContext } from '../../src/formats/pptx/writer/text-node';

function context(...fieldIds: string[]) {
  let index = 0;
  return {
    allocateFieldId: vi.fn(() => {
      const fieldId = fieldIds[index];
      index += 1;
      return fieldId ?? '{missing}';
    }),
  } satisfies PptxTextSerializationContext;
}

describe('PowerPoint text-body serialization', () => {
  it('serializes an empty property set without invented attributes', () => {
    expect(serializeTextBodyProperties({})).toBe('<a:bodyPr/>');
  });

  it.each([
    ['top', 't'],
    ['center', 'ctr'],
    ['bottom', 'b'],
    ['justified', 'just'],
    ['distributed', 'dist'],
  ] as const)('maps %s anchor to %s', (anchor, expected) => {
    expect(serializeTextBodyProperties({ anchor })).toBe(
      `<a:bodyPr anchor="${expected}"/>`,
    );
  });

  it('distinguishes explicit horizontal and vertical text', () => {
    expect(serializeTextBodyProperties({ vertical: true })).toBe(
      '<a:bodyPr vert="eaVert"/>',
    );
    expect(serializeTextBodyProperties({ vertical: false })).toBe(
      '<a:bodyPr vert="horz"/>',
    );
  });

  it('distinguishes explicit wrapping modes', () => {
    expect(serializeTextBodyProperties({ wrap: true })).toBe(
      '<a:bodyPr wrap="square"/>',
    );
    expect(serializeTextBodyProperties({ wrap: false })).toBe(
      '<a:bodyPr wrap="none"/>',
    );
  });

  it.each([
    ['none', '<a:noAutofit/>'],
    ['shape', '<a:spAutoFit/>'],
    ['text', '<a:normAutofit/>'],
  ] as const)(
    'maps %s auto-fit to its DrawingML child',
    (autoFit, expected) => {
      expect(serializeTextBodyProperties({ autoFit })).toBe(
        `<a:bodyPr>${expected}</a:bodyPr>`,
      );
    },
  );

  it('serializes attributes and auto-fit in deterministic order', () => {
    expect(
      serializeTextBodyProperties({
        anchor: 'center',
        autoFit: 'shape',
        vertical: false,
        wrap: true,
      }),
    ).toBe(
      '<a:bodyPr anchor="ctr" vert="horz" wrap="square"><a:spAutoFit/></a:bodyPr>',
    );
  });

  it('serializes list style before every ordered paragraph', () => {
    const allocation = context('{field-1}');
    const text: PptxSceneTextBody = {
      body: { anchor: 'top', autoFit: 'none' },
      paragraphs: [
        {
          children: [{ key: 'run-1', text: 'First', type: 'run' }],
          key: 'paragraph-1',
        },
        {
          children: [
            {
              fieldType: 'slidenum',
              key: 'field-1',
              text: '2',
              type: 'field',
            },
          ],
          key: 'paragraph-2',
          properties: { alignment: 'right', level: 1 },
        },
      ],
    };

    expect(serializeTextBody(text, allocation)).toBe(
      '<p:txBody><a:bodyPr anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr/><a:t>First</a:t></a:r></a:p><a:p><a:pPr lvl="1" algn="r"/><a:fld id="{field-1}" type="slidenum"><a:rPr/><a:t>2</a:t></a:fld></a:p></p:txBody>',
    );
    expect(allocation.allocateFieldId).toHaveBeenCalledTimes(1);
  });
});
