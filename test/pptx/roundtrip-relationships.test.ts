import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { resolvePptxSlideParts } from '../../src/formats/pptx/roundtrip/relationships';
import { PptxWriteError } from '../../src/formats/pptx/write-error';
import { createMinimalPptx } from './fixture';

const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

async function archive(
  overrides: Record<string, string | null> = {},
): Promise<JSZip> {
  return JSZip.loadAsync(await createMinimalPptx(overrides));
}

describe('PowerPoint editable slide relationship resolution', () => {
  it('follows presentation order across two internal slide relationships', async () => {
    const presentation = `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`;
    const relationships = `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/slide2.xml"/></Relationships>`;
    const value = await archive({
      'ppt/_rels/presentation.xml.rels': relationships,
      'ppt/presentation.xml': presentation,
      'ppt/slides/slide2.xml': `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    });

    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).resolves.toEqual(['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml']);
  });

  it('ignores external, non-slide, and incomplete relationships', async () => {
    const relationships = `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">
      <Relationship Id="external" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="https://example.test/deck" TargetMode="External"/>
      <Relationship Id="wrong" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/notesSlide" Target="notesSlides/notesSlide1.xml"/>
      <Relationship Id="missingTarget" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide"/>
      <Relationship Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/orphan.xml"/>
      <Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/slide1.xml"/>
    </Relationships>`;
    const value = await archive({
      'ppt/_rels/presentation.xml.rels': relationships,
    });

    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).resolves.toEqual(['ppt/slides/slide1.xml']);
  });

  it.each(['External', 'EXTERNAL', 'external'])(
    'does not resolve the selected slide when TargetMode=%s',
    async (mode) => {
      const value = await archive({
        'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/slide1.xml" TargetMode="${mode}"/></Relationships>`,
      });

      await expect(
        resolvePptxSlideParts(value, resolvePptxResourceLimits()),
      ).rejects.toThrow(
        'PowerPoint text edit cannot resolve the presentation slide order',
      );
    },
  );

  it.each([
    [
      'unsafe target',
      `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="../../../evil.xml"/></Relationships>`,
      'PowerPoint text edit encountered an unsafe slide relationship',
    ],
    [
      'missing target part',
      `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/slide" Target="slides/missing.xml"/></Relationships>`,
      'PowerPoint text edit cannot resolve the presentation slide order',
    ],
  ])('rejects %s', async (_name, relationships, message) => {
    const value = await archive({
      'ppt/_rels/presentation.xml.rels': relationships,
    });
    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).rejects.toThrow(message);
  });

  it('does not select a relationship with a non-slide type', async () => {
    const value = await archive({
      'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdSlide1" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/notesSlide" Target="slides/slide1.xml"/></Relationships>`,
    });
    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).rejects.toThrow(
      'PowerPoint text edit cannot resolve the presentation slide order',
    );
  });

  it('rejects a slide id without a relationship id', async () => {
    const value = await archive({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
    });
    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).rejects.toThrow(
      'PowerPoint text edit cannot resolve the presentation slide order',
    );
  });

  it('accepts an empty presentation without inventing slide entries', async () => {
    const value = await archive({
      'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"/>`,
      'ppt/presentation.xml': `<p:presentation xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"/>`,
    });

    await expect(
      resolvePptxSlideParts(value, resolvePptxResourceLimits()),
    ).resolves.toEqual([]);
  });

  it('rejects missing, malformed, and oversized required XML parts', async () => {
    const missing = await archive({ 'ppt/presentation.xml': null });
    await expect(
      resolvePptxSlideParts(missing, resolvePptxResourceLimits()),
    ).rejects.toThrow(
      'PowerPoint text edit cannot read required part ppt/presentation.xml',
    );

    const malformed = await archive({
      'ppt/_rels/presentation.xml.rels': '<Relationships>',
    });
    const malformedResult = resolvePptxSlideParts(
      malformed,
      resolvePptxResourceLimits(),
    );
    await expect(malformedResult).rejects.toThrow(
      'PowerPoint text edit cannot read required part ppt/_rels/presentation.xml.rels',
    );
    const malformedError: unknown = await malformedResult.catch(
      (error: unknown) => error,
    );
    expect(malformedError).toBeInstanceOf(PptxWriteError);
    if (!(malformedError instanceof PptxWriteError)) throw malformedError;
    expect(malformedError.cause).toBeInstanceOf(Error);

    const limits = resolvePptxResourceLimits();
    limits.maxXmlBytes = 8;
    await expect(
      resolvePptxSlideParts(await archive(), limits),
    ).rejects.toThrow('PowerPoint text edit cannot read required part');
  });
});
