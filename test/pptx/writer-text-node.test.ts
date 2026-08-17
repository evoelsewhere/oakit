import { describe, expect, it, vi } from 'vitest';

import type {
  PptxSceneRunProperties,
  PptxSceneTextNode,
} from '../../src/formats/pptx/scene-types';
import {
  serializeEndParagraphProperties,
  serializeRunProperties,
  serializeTextNode,
  type PptxTextSerializationContext,
} from '../../src/formats/pptx/writer/text-node';

function context(fieldId = '{00000000-0000-0000-0000-000000000001}') {
  return {
    allocateFieldId: vi.fn(() => fieldId),
  } satisfies PptxTextSerializationContext;
}

describe('PowerPoint text-node serialization', () => {
  it('serializes absent and empty run properties identically', () => {
    expect(serializeRunProperties()).toBe('<a:rPr/>');
    expect(serializeRunProperties({})).toBe('<a:rPr/>');
    expect(serializeEndParagraphProperties()).toBe('<a:endParaRPr/>');
  });

  it('serializes every run attribute in deterministic order', () => {
    const properties: PptxSceneRunProperties = {
      bold: true,
      fontSize: 18.006,
      italic: false,
      language: 'en-US',
    };

    expect(serializeRunProperties(properties)).toBe(
      '<a:rPr lang="en-US" sz="1801" b="1" i="0"/>',
    );
  });

  it('distinguishes explicit false and true boolean properties', () => {
    expect(serializeRunProperties({ bold: false, italic: true })).toBe(
      '<a:rPr b="0" i="1"/>',
    );
  });

  it('applies one escaped font family to every script class', () => {
    expect(serializeRunProperties({ fontFamily: `A & "B" <C>` })).toBe(
      '<a:rPr><a:latin typeface="A &amp; &quot;B&quot; &lt;C&gt;"/><a:ea typeface="A &amp; &quot;B&quot; &lt;C&gt;"/><a:cs typeface="A &amp; &quot;B&quot; &lt;C&gt;"/></a:rPr>',
    );
  });

  it('serializes run color before font script mappings', () => {
    expect(
      serializeRunProperties({ color: '#f8FaFc', fontFamily: 'Aptos' }),
    ).toBe(
      '<a:rPr><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill><a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/></a:rPr>',
    );
  });

  it('uses the end-paragraph element without changing its properties', () => {
    expect(
      serializeEndParagraphProperties({
        bold: true,
        color: '#112233',
        fontFamily: 'Aptos',
        fontSize: 12,
        italic: true,
        language: 'vi-VN',
      }),
    ).toBe(
      '<a:endParaRPr lang="vi-VN" sz="1200" b="1" i="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/></a:endParaRPr>',
    );
  });

  it('serializes a run without allocating a field identity', () => {
    const allocation = context();
    const node: PptxSceneTextNode = {
      key: 'run-1',
      preserveSpace: true,
      properties: { bold: true },
      text: '<&_x0041_',
      type: 'run',
    };

    expect(serializeTextNode(node, allocation)).toBe(
      '<a:r><a:rPr b="1"/><a:t xml:space="preserve">&lt;&amp;_x005F_x0041_</a:t></a:r>',
    );
    expect(allocation.allocateFieldId).not.toHaveBeenCalled();
  });

  it('does not force preserved spacing when a run opts out', () => {
    const node: PptxSceneTextNode = {
      key: 'run-1',
      preserveSpace: false,
      text: 'plain',
      type: 'run',
    };

    expect(serializeTextNode(node, context())).toBe(
      '<a:r><a:rPr/><a:t>plain</a:t></a:r>',
    );
  });

  it('still preserves whitespace required by the text value', () => {
    const node: PptxSceneTextNode = {
      key: 'run-1',
      text: ' leading',
      type: 'run',
    };

    expect(serializeTextNode(node, context())).toBe(
      '<a:r><a:rPr/><a:t xml:space="preserve"> leading</a:t></a:r>',
    );
  });

  it('serializes fields with one allocated and escaped identity', () => {
    const allocation = context(`{id&"}`);
    const node: PptxSceneTextNode = {
      fieldType: `slide<&"`,
      key: 'field-1',
      properties: { italic: true },
      text: ' 2 ',
      type: 'field',
    };

    expect(serializeTextNode(node, allocation)).toBe(
      '<a:fld id="{id&amp;&quot;}" type="slide&lt;&amp;&quot;"><a:rPr i="1"/><a:t xml:space="preserve"> 2 </a:t></a:fld>',
    );
    expect(allocation.allocateFieldId).toHaveBeenCalledTimes(1);
  });

  it('serializes breaks without allocating a field identity', () => {
    const allocation = context();
    const node: PptxSceneTextNode = {
      key: 'break-1',
      properties: { language: 'vi-VN' },
      type: 'break',
    };

    expect(serializeTextNode(node, allocation)).toBe(
      '<a:br><a:rPr lang="vi-VN"/></a:br>',
    );
    expect(allocation.allocateFieldId).not.toHaveBeenCalled();
  });
});
