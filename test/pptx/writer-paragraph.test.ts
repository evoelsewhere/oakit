import { describe, expect, it, vi } from 'vitest';

import type { PptxSceneParagraph } from '../../src/formats/pptx/scene-types';
import {
  serializeParagraph,
  serializeParagraphProperties,
} from '../../src/formats/pptx/writer/paragraph';
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

describe('PowerPoint paragraph serialization', () => {
  it('preserves absent paragraph properties', () => {
    expect(serializeParagraphProperties(undefined)).toBe('');
    expect(
      serializeParagraph({ children: [], key: 'paragraph-1' }, context()),
    ).toBe('<a:p></a:p>');
  });

  it('preserves an explicitly empty paragraph property object', () => {
    expect(serializeParagraphProperties({})).toBe('<a:pPr/>');
  });

  it.each([
    ['left', 'l'],
    ['center', 'ctr'],
    ['right', 'r'],
    ['justify', 'just'],
    ['distributed', 'dist'],
  ] as const)('maps %s alignment to %s', (alignment, expected) => {
    expect(serializeParagraphProperties({ alignment })).toBe(
      `<a:pPr algn="${expected}"/>`,
    );
  });

  it.each([0, 1, 8])('serializes paragraph level %s exactly', (level) => {
    expect(serializeParagraphProperties({ level })).toBe(
      `<a:pPr lvl="${level}"/>`,
    );
  });

  it('serializes level before alignment deterministically', () => {
    expect(
      serializeParagraphProperties({ alignment: 'center', level: 2 }),
    ).toBe('<a:pPr lvl="2" algn="ctr"/>');
  });

  it('keeps structured children in authored order', () => {
    const allocation = context('{field-1}', '{field-2}');
    const paragraph: PptxSceneParagraph = {
      children: [
        { key: 'run-1', text: 'A', type: 'run' },
        {
          fieldType: 'slidenum',
          key: 'field-1',
          text: '1',
          type: 'field',
        },
        { key: 'break-1', type: 'break' },
        {
          fieldType: 'datetime',
          key: 'field-2',
          text: 'Now',
          type: 'field',
        },
        { key: 'run-2', text: 'B', type: 'run' },
      ],
      key: 'paragraph-1',
      properties: { alignment: 'left', level: 0 },
    };

    expect(serializeParagraph(paragraph, allocation)).toBe(
      '<a:p><a:pPr lvl="0" algn="l"/><a:r><a:rPr/><a:t>A</a:t></a:r><a:fld id="{field-1}" type="slidenum"><a:rPr/><a:t>1</a:t></a:fld><a:br><a:rPr/></a:br><a:fld id="{field-2}" type="datetime"><a:rPr/><a:t>Now</a:t></a:fld><a:r><a:rPr/><a:t>B</a:t></a:r></a:p>',
    );
    expect(allocation.allocateFieldId).toHaveBeenCalledTimes(2);
  });

  it('places end-paragraph properties after every child', () => {
    const paragraph: PptxSceneParagraph = {
      children: [{ key: 'run-1', text: 'Text', type: 'run' }],
      endProperties: {
        bold: false,
        fontFamily: 'Aptos',
        fontSize: 11,
        italic: true,
        language: 'vi-VN',
      },
      key: 'paragraph-1',
    };

    expect(serializeParagraph(paragraph, context())).toBe(
      '<a:p><a:r><a:rPr/><a:t>Text</a:t></a:r><a:endParaRPr lang="vi-VN" sz="1100" b="0" i="1"><a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/></a:endParaRPr></a:p>',
    );
  });

  it('preserves an explicitly empty end-paragraph property object', () => {
    const paragraph: PptxSceneParagraph = {
      children: [],
      endProperties: {},
      key: 'paragraph-1',
    };

    expect(serializeParagraph(paragraph, context())).toBe(
      '<a:p><a:endParaRPr/></a:p>',
    );
  });
});
