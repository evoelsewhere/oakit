# PPTX Parser-to-Writer and Round-Trip Plan

> Status: active foundation plan. The frozen production capability remains
> PPTX-to-JSON reading until individual creation and edit capability rows pass
> their gates. This document does not claim that unfinished JSON-to-PPTX or
> edited round-tripping is implemented.

This plan defines how OAKit should evolve the PowerPoint parser model, add a
strict JSON-to-PPTX writer, and add a source-preserving edit pipeline. The work
starts only after the currently active goal is complete and its required gates
are green.

The central architectural decision is that there are two distinct write cases:

1. **Create from semantic JSON:** generate a new, valid presentation from an
   OAKit-owned document model.
2. **Edit an existing presentation:** preserve the original package, change a
   declared set of supported features, and prove what remained unchanged.

The current renderer-oriented `PptxDocument` is not sufficient as the only
input to a high-fidelity writer. It contains useful geometry and visual data,
but parsing currently resolves or flattens information that PowerPoint needs
when authoring. A reversible design therefore requires a richer semantic model
and a separate opaque preservation envelope.

## Kickoff audit and first delivery boundary

The reader reliability baseline was frozen at commit `64ae733` on 2026-08-16.
At kickoff it passes formatting, lint, strict type checking, 1,986 unit and
property tests, 14 producer-corpus cases, 15 browser tests across Chromium,
Firefox, and WebKit, package smoke tests, and the complete mutation audit with
7,409 killed mutants, 2,193 compile errors, and zero missed mutants across 38
runtime files.

The architecture below remains the long-term target, but fidelity levels are
awarded per versioned capability row rather than to the entire PPTX format.
The first implementation slice is deliberately narrow:

1. freeze additive experimental V2 scene types, stable keys, validation, and a
   support profile without changing the existing V1 reader contract;
2. prove C1 creation for deterministic minimal presentations and structured
   text before adding other element domains;
3. prove R0 exact output for every accepted source, independently through
   runtime bytes/Blob and portable JSON snapshots;
4. prove R1/R2 first for a bounded text operation whose dirty closure is one
   known slide part, while preserving every copied part payload exactly;
5. keep every unimplemented subtype, owner, conformance mode, and operation
   explicitly unsupported or preservation-only.

This ordering does not advertise general PowerPoint writing. It establishes
the package, security, consistency, verification, and mutation infrastructure
once, then advances one capability row at a time.

The empty-hierarchy contract is also made explicit: a source-free slide may
omit `layoutKey` only when `themes`, `masters`, and `layouts` are all empty and
the writer is asked to allocate the documented minimal hierarchy. If any
hierarchy object is supplied, every slide-to-layout, layout-to-master, and
master-to-theme reference is required and validated; contradictory partial
hierarchies are rejected.

## Direct answer: what “100%” can mean

OOXML and producer-specific PowerPoint extensions are open-ended. “100% PPTX
round-trip” must be attached to a measurable guarantee, not used as a blanket
claim.

| Guarantee                                                                                       | Achievable                               | Required design                                                                                                                              |
| ----------------------------------------------------------------------------------------------- | ---------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| Return exactly the original bytes when no edit is requested                                     | Yes                                      | Preserve complete source bytes in a runtime snapshot or bounded portable Base64 form, verify the hash, and return without rebuilding the ZIP |
| Create a valid PPTX from every field in a declared semantic schema version                      | Yes                                      | Strict model validation, a complete mapping matrix, deterministic serializers, and fresh-reader verification                                 |
| Preserve every untouched source part payload during a supported edit                            | Yes, within a declared operation profile | Copy each untouched part's uncompressed bytes exactly and calculate the complete dirty dependency closure before writing                     |
| Preserve semantics of every supported edited feature                                            | Yes, per feature and operation           | Independent fixtures, output reparse, package graph validation, rendering checks, and producer verification                                  |
| Reconstruct an arbitrary original deck exactly from normalized JSON alone                       | No                                       | The normalized model deliberately removes relationship IDs, source ownership, unknown XML, ZIP metadata, and producer extensions             |
| Safely edit every future or unknown PowerPoint feature                                          | No                                       | Unknown content may be preserved only when proven independent of a dirty part; otherwise the edit must be blocked                            |
| Guarantee pixel-identical rendering on every PowerPoint, LibreOffice, and Google Slides version | No                                       | Rendering depends on fonts, producer versions, operating systems, fallback behavior, and undocumented extensions                             |

The public promise should therefore be:

> OAKit creates package-valid presentations from the supported semantic model,
> returns the exact source bytes for an unchanged round-trip snapshot, and
> preserves untouched content while applying only edits covered by a versioned
> capability profile. It blocks an edit when that fidelity cannot be proven.

## Outcome of the future goal

The future goal is complete only when all three deliverables exist:

1. a round-trip-aware PowerPoint semantic model and parser;
2. a source-free JSON-to-PPTX creation writer;
3. a source-preserving PPTX edit writer with explicit fidelity evidence.

These deliverables share domain types and serialization rules, but the reader,
creation writer, and round-trip editor remain separate orchestrators.

```text
                         semantic scene model
                        /                    \
PPTX -> strict reader -+                      +-> creation writer -> new PPTX
                        \                    /
                         round-trip snapshot
                           + edit operations
                                  |
                                  v
                        preservation writer -> edited PPTX
```

## Fidelity levels

Every write result reports one of the following levels. The level is computed
from evidence; callers cannot request a name and thereby force the writer to
claim it.

| Level | Meaning                                                                                                                                                                                |
| ----- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `C1`  | Source-free creation. The output is package-valid and reparses to the normalized supported semantic input.                                                                             |
| `C2`  | Verified creation. In addition to C1, rendering and producer-open checks pass for every feature used by the document.                                                                  |
| `R0`  | Exact no-op. There are no edits and the returned bytes are identical to the verified source bytes.                                                                                     |
| `R1`  | Part-preserving edit. Every untouched part's uncompressed payload is copied exactly; all dirty parts belong to supported operation paths and pass graph validation.                    |
| `R2`  | Semantic edit verification. In addition to R1, a fresh strict parse proves the requested semantic changes and required invariants.                                                     |
| `R3`  | Producer-verified edit. In addition to R2, the output opens without repair and preserves declared behavior in the supported PowerPoint, LibreOffice Impress, and Google Slides matrix. |

`R0` is an exact terminal class, not a lower-quality edited output. `C1` and
`C2` apply only to newly created packages and must never be described as source
preservation.

The initial implementation may ship one supported domain at a time. General
PPTX editing must not be advertised until the public support matrix and the
actual evidence agree.

## Architectural principles

### Preserve the master-layout-slide hierarchy

PowerPoint authoring is based on this ownership chain:

```text
theme -> slide master -> slide layout -> slide
```

The model must preserve these owners rather than flattening all inherited
objects into a slide. A slide selects a layout; the layout selects a master;
placeholders link the objects across those levels. Resolved properties may be
provided for convenient rendering, but the authored owner and inheritance
source remain available to the writer.

This is required to prevent common reverse-mapping bugs:

- writing the same logo or footer into every slide instead of its master;
- converting a placeholder into an unrelated freeform shape;
- duplicating inherited title or body properties on every slide;
- changing one slide when the requested edit should affect a layout;
- changing all slides when the edit should create a local override;
- losing `showMasterSp`, color-map overrides, or layout visibility behavior.

### Ordered element arrays are the z-order contract

Every shape-tree owner exposes one canonical ordered array. Earlier entries are
behind later entries. Nested groups own their own ordered child arrays.

The model must not store pairwise claims such as `aboveElementId` or
`belowElementId`. Pairwise relationships become contradictory after insertion,
deletion, grouping, or reordering. Stable operations express intent instead:

- `moveElementBefore`;
- `moveElementAfter`;
- `sendElementToBack`;
- `bringElementToFront`;
- `moveElementToGroup`.

The compatibility `order` field may be derived from array position. It is not
the canonical V2 source of truth.

### Coordinates are necessary but must identify their space

Every rendered visual element needs a resolved position, size, rotation, and
flips. The authored element may omit its transform and inherit it from a
placeholder ancestor; that absence remains explicit. Coordinates use points at
the public boundary, and the model identifies which owner or group coordinate
space authored and resolved transforms belong to.

Groups require both the outer transform and their child coordinate space. A
writer cannot correctly reconstruct nested group transforms from flattened
absolute child positions alone.

```ts
interface PptxTransform {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

interface PptxGroupTransform extends PptxTransform {
  childSpace: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
}
```

For a source-free document, point values are converted deterministically to
integer EMUs. For a source-preserving edit, unchanged transforms remain in the
copied source part; edited values are converted through one checked unit
serializer.

### Semantic data and preservation data are separate

Normalized public data must not contain raw XML, relationship IDs, ZIP entry
objects, parser caches, or mutable internal nodes. Exact source preservation is
instead carried in a versioned round-trip envelope:

```text
round-trip snapshot
├── semantic PowerPoint scene model
├── stable OAKit-owned keys
├── versioned support profile
├── ordered typed operations
├── opaque source bytes or bounded portable Base64
├── source SHA-256 and consistency metadata
└── bounded diagnostics and fidelity evidence
```

The source package is untrusted binary data and is used only by the strict
round-trip pipeline. It is never treated as semantic JSON and never exposed as
raw XML fields on slide or element objects.

### Operations, not arbitrary JSON diffs, drive edits

Directly mutating a parsed semantic object does not reveal whether a change
should patch a slide, layout, master, theme, relationship, chart cache, media
part, notes part, or several of them. It also cannot express preconditions or
conflicts with unknown extensions.

The source-preserving API therefore accepts typed, ordered operations. The
semantic document inside a snapshot is a preview and targeting surface; the
operation list is the authoritative edit intent.

### Writer behavior is strict

The current reader may recover from documented optional failures. The writer
must not emit a partial or guessed package. A write either:

- validates and produces complete bytes plus a fidelity report; or
- throws a typed error before returning output.

There is no tolerant write mode and no `bestEffort: true` escape hatch.

### Reader, creation writer, and preservation writer stay separate

The creation writer is not implemented by running parser functions backward.
Reader internals resolve inheritance and normalize producer variation; writer
internals allocate identifiers and serialize canonical OOXML. Only truly
bidirectional domain contracts are shared.

## Audit of the current public model

The existing model already provides useful creation inputs:

- presentation width and height;
- per-element position and size in points;
- rotation and flips for most visual objects;
- element `order` and nested groups;
- slide-authored versus inherited `layoutElements`;
- resolved fills, borders, shadows, image crops, and filters;
- rendered text HTML;
- normalized tables, charts, diagrams, math, notes, and transitions.

The following losses prevent reliable reverse mapping from that model alone.

| Domain              | Current normalization                                                       | Information required for authoring                                                                                                                                 |
| ------------------- | --------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Master/layout/slide | Inheritance is resolved and decorations are flattened into `layoutElements` | Explicit masters, layouts, slide ownership, hierarchy keys, visibility flags, placeholder links, and local overrides                                               |
| Text                | Paragraphs and runs become HTML                                             | Ordered paragraph/run/field/break nodes, language, fonts, theme references, bullets, numbering, tabs, spacing, hyperlinks, autofit, and inherited property origins |
| Geometry            | Positions are converted to points; group children are transformed           | Coordinate-space identity, original group `off/ext/chOff/chExt`, connector endpoints, connection sites, adjustment guides, and custom-geometry commands            |
| Z-order             | Numeric `order` survives                                                    | One canonical ordered array at every owner and group, plus stable edit operations                                                                                  |
| Color               | Colors usually become resolved hex values                                   | Color kind, scheme/system/preset source, color-map owner, ordered transforms, opacity, and optional resolved preview                                               |
| Fill/line/effect    | Renderer-oriented values                                                    | Full authoring variants, cap/join/compound/dash, gradient stops and geometry, pattern identity, effect stack, and inheritance origin                               |
| Media               | `ref`, Base64, or object URL strings                                        | JSON-stable representation groups, MIME type, bytes or external target, fallback/poster roles, crop, embedding policy, and relationship ownership                  |
| Chart               | Caches become normalized series                                             | ChartSpace structure, formulas and caches separately, axes, titles, labels, legend, formats, external-data relationship, and embedded workbook ownership           |
| Table               | Cell presentation is mostly flattened                                       | Structured rich text, table style identity, row/column properties, margins, merges, banding flags, and authored versus resolved style                              |
| SmartArt            | Exposed as flattened shapes/text plus a text list                           | Data model, layout/style/color references, drawing part ownership, and an opaque-preservation boundary for unsupported algorithms                                  |
| Notes               | Body becomes HTML                                                           | Structured notes text, notes-master inheritance, placeholders, notes size, and notes relationships                                                                 |
| Animation           | Slide transition is modeled                                                 | Main sequence, timing tree, targets, triggers, conditions, media timing, and unsupported extension policy                                                          |
| Package metadata    | Mostly absent                                                               | Document/core/custom properties, sections, custom shows, comments, tags, embedded fonts, extension lists, and conformance mode                                     |

The V1 model remains valuable for current consumers. The migration must be
deliberate rather than silently changing the meaning of existing fields.

## Target semantic model

Public names in this section are provisional until contract tests freeze them.
The shape of the model is normative: authored owners, ordered trees, stable
keys, structured text, and theme-aware values are required.

### Document graph

```ts
interface PptxSceneDocument {
  schemaVersion: 2;
  size: { width: number; height: number };
  notesSize?: { width: number; height: number };
  properties?: PptxDocumentProperties;
  presentationDefaults?: PptxPresentationDefaults;
  themes: PptxTheme[];
  masters: PptxSlideMaster[];
  layouts: PptxSlideLayout[];
  notesMasters?: PptxNotesMaster[];
  handoutMasters?: PptxHandoutMaster[];
  slides: PptxSceneSlide[];
  media: PptxMediaResource[];
  embeddedFonts?: PptxEmbeddedFont[];
  sections?: PptxSection[];
  customShows?: PptxCustomShow[];
  comments?: PptxCommentStore;
}

interface PptxSlideMaster {
  key: string;
  themeKey: string;
  name?: string;
  background?: PptxBackground;
  elements: PptxSceneElement[];
  textStyles?: PptxMasterTextStyles;
  colorMap?: PptxColorMap;
  preserve?: boolean;
}

interface PptxSlideLayout {
  key: string;
  masterKey: string;
  name?: string;
  layoutType?: string;
  matchingName?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  background?: PptxBackground;
  elements: PptxSceneElement[];
  preserve?: boolean;
  userDrawn?: boolean;
}

interface PptxSceneSlide {
  key: string;
  layoutKey?: string;
  name?: string;
  hidden?: boolean;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  background?: PptxBackground;
  elements: PptxSceneElement[];
  notes?: PptxNotes;
  transition?: PptxTransition;
  animations?: PptxAnimationSequence;
}
```

Keys are assigned by OAKit and are stable within one snapshot. They are not
OOXML relationship IDs, part filenames, or shape IDs. A source-free document
may use caller-provided unique keys; the writer validates them and allocates all
package-level identifiers independently. The canonical scene returned by the
parser always contains normalized hierarchy arrays. For source-free creation,
empty `themes`, `masters`, or `layouts` arrays request the documented minimal
hierarchy; omitted hierarchy is not represented through contradictory partial
objects.

### Element base and ownership

Every element contains:

- a stable `key`;
- an optional author-facing name and alt text;
- authored properties whose absence remains meaningful;
- a resolved preview for rendering and V1 compatibility;
- visibility and locking properties where supported;
- a discriminated semantic payload;
- optional placeholder linkage;
- optional hyperlink/action data;
- bounded value-source metadata explaining resolved inheritance without
  exposing raw OOXML.

```ts
interface PptxElementBase {
  key: string;
  name?: string;
  description?: string;
  title?: string;
  decorative?: boolean;
  authored: {
    transform?: PptxTransform;
    hidden?: boolean;
    locks?: PptxElementLocks;
  };
  resolved: {
    transform?: PptxTransform;
    hidden: boolean;
  };
  valueSources?: PptxElementValueSources;
  placeholder?: PptxPlaceholderMetadata;
}

interface PptxPlaceholderMetadata {
  type?: string;
  index?: number;
  orientation?: 'horizontal' | 'vertical';
  size?: 'full' | 'half' | 'quarter';
  hasCustomPrompt?: boolean;
  prompt?: string;
  sourceKey?: string;
  role: 'master-definition' | 'layout-definition' | 'slide-instance';
}
```

The authored/resolved split applies to every property that can inherit, not
only transforms. Text, geometry, fills, lines, effects, colors, backgrounds,
and placeholder properties retain an authored container, a resolved preview,
and bounded value-source metadata. The writer consumes authored state. It must
never serialize a resolved fallback merely because it exists in the preview.

An absent authored value is different from explicit `null`, zero, `false`, an
empty collection, or a default-valued OOXML attribute. The V2 validator and
JSON schema preserve this distinction.

### Placeholder ownership and composite rendering

A placeholder is an element with one canonical `placeholder` metadata object.
Masters and layouts do not carry a second independently editable placeholder
array. Parser indexes used for matching by source ID, `idx`, type, orientation,
or size are derived internal data and cannot conflict with `elements`.

Placeholder inheritance follows explicit source keys where present and
deterministic PowerPoint-compatible matching otherwise. The resolved
composition algorithm is normative:

```text
slide background resolution
  -> visible master-owned decorative elements
  -> visible layout-owned decorative elements
  -> slide-owned elements in authored shape-tree order
```

Placeholder definitions provide inherited properties but are not rendered as
additional duplicate objects when a descendant instance represents them.
`showMasterSp`, `showMasterPhAnim`, layout flags, placeholder visibility, and
local overrides participate in composition. Editing a master, layout, theme,
or placeholder requires rendering representative descendant slides during
verification.

The element union contains at least:

```text
text | shape | connector | image | table | chart | group
audio | video | diagram | math | oleObject | unsupported
```

`oleObject` is preservation-only at first. `unsupported` may appear only in a
round-trip snapshot with bounded metadata; it is never accepted by the
source-free creation writer unless an explicit future serializer supports it.

### Structured text is the source of truth

Rendered HTML remains available as a derived compatibility view, not as the
writer input.

```ts
interface PptxTextBody {
  body: PptxTextBodyProperties;
  paragraphs: PptxParagraph[];
}

interface PptxParagraph {
  key: string;
  properties?: PptxParagraphProperties;
  children: Array<PptxTextRun | PptxTextField | PptxTextBreak>;
  endProperties?: PptxRunProperties;
}

interface PptxTextRun {
  type: 'run';
  key: string;
  text: string;
  preserveSpace?: boolean;
  properties?: PptxRunProperties;
}

interface PptxTextField {
  type: 'field';
  key: string;
  fieldType: string;
  text: string;
  properties?: PptxRunProperties;
}

interface PptxTextBreak {
  type: 'break';
  key: string;
  properties?: PptxRunProperties;
}
```

Paragraph properties cover alignment, margin level, indentation, tabs,
line/space-before/space-after, direction, bullet/numbering, default run
properties, and inherited source. Run properties cover typeface, theme font,
size, bold, italic, underline, strike, baseline, capitalization, language,
character spacing, color, highlight, hyperlink/action, and supported effects.

The parser must retain field identity for slide numbers, date/time, and footer
fields. Converting fields to plain text would make a reverse mapping incorrect.
It must also distinguish authored tabs, line breaks, significant leading or
trailing spaces, literal strings resembling Office `_xHHHH_` escapes, and
characters that XML 1.0 cannot represent directly. The writer owns one tested
Office-text escaping policy and emits `xml:space="preserve"` when required.

### Theme-aware color model

```ts
type PptxColor =
  | { type: 'rgb'; value: string; transforms?: PptxColorTransform[] }
  | { type: 'scheme'; value: string; transforms?: PptxColorTransform[] }
  | {
      type: 'system';
      value: string;
      lastColor?: string;
      transforms?: PptxColorTransform[];
    }
  | { type: 'preset'; value: string; transforms?: PptxColorTransform[] };
```

An optional `resolved` preview may be cached for consumers, but the authored
color and ordered transforms are canonical. This allows theme changes to remain
theme changes instead of being expanded into fixed RGB values.

### Shape, connector, and group geometry

Shape geometry is a union:

- preset geometry plus adjustment values;
- custom geometry with guides, handles, connection sites, text rectangle, and
  ordered path commands;
- preservation-only geometry for an unsupported extension.

SVG path and view-box fields are derived renderer output. They are not the
authoring source because the conversion loses guide formulas and connection
semantics.

Connectors are a distinct element type. Each endpoint may reference a stable
element key and connection-site index, with a fallback point for unresolved or
source-free connectors.

Groups preserve their outer transform and child coordinate space. Children
remain local to that group; the V2 parser must stop making flattened absolute
coordinates the only representation.

### Media registry

Binary assets live once in a document-level registry and elements reference a
media key.

```ts
interface PptxMediaResource {
  key: string;
  kind: 'image' | 'audio' | 'video' | 'font' | 'ole' | 'other';
  primaryRepresentationKey: string;
  representations: PptxMediaRepresentation[];
  usageRoles?: PptxMediaUsageRole[];
}

interface PptxMediaRepresentation {
  key: string;
  role: 'primary' | 'fallback' | 'poster' | 'preview' | 'payload';
  mimeType: string;
  source:
    | { type: 'embedded'; base64: string }
    | { type: 'snapshot'; sourceResourceKey: string; sha256: string }
    | { type: 'external'; target: string };
  sha256?: string;
  fileNameHint?: string;
}
```

Object URLs are not valid JSON authoring inputs. The existing object-URL modes
remain reader compatibility options, but the scene/round-trip contract uses
embedded Base64, an OAKit-owned round-trip source-resource key, or an explicit
external target. No external target is fetched. Source-free creation rejects
`snapshot` media sources because it has no opaque package from which to resolve
them.

A round-trip scene uses `snapshot` references for unchanged embedded source
media instead of copying every payload into semantic JSON as well as the opaque
source. An operation that adds or replaces media supplies bounded embedded
Base64 or a separately authorized caller-provided byte input through a future
binary operation API.

Representation groups preserve cases such as SVG plus PNG fallback, EMF/WMF
plus raster preview, and video plus poster image. Source-preserving mode keeps
distinct source media parts distinct by default even when payload hashes are
equal. Source-free creation may deduplicate only when all relationship roles,
fallback behavior, metadata, and producer semantics are equivalent.

Embedded fonts are not ordinary media. Their model and policy retain
obfuscation identity, subset/full status, relationship ownership, and embedding
rights. OAKit does not extract, de-obfuscate, rewrite, or redistribute a font
unless the operation and declared rights permit it.

### Charts, tables, diagrams, notes, and animation

- A chart separates formulas, cached values, displayed formatting, axes,
  legend, labels, plot area, and optional embedded workbook references.
- A table stores rich text per cell, row/column sizing, merges, margins,
  authored fills/borders, and the table-style identity separately from resolved
  preview values.
- A diagram keeps logical SmartArt data and part-role references when
  understood. Unsupported layout algorithms are preservation-only.
- Office Math uses a tested semantic math AST when authoring is supported.
  LaTeX and fallback images are derived views; LaTeX alone is never considered
  a reversible OMML source. Unsupported OMML remains preservation-only.
- Notes use the same structured text primitives and retain notes-slide and
  notes-master ownership. PowerPoint has no slide-layout-equivalent notes part.
- Legacy comments/authors and modern comments/people/threading are separate
  versioned models with explicit conversion policy.
- Slide transitions remain separate from the full timing/animation tree.
  Animation writing is not advertised until targets, timing, and triggers have
  independent verification.

## Parser migration strategy

### Keep V1 stable while V2 is built

The existing `parsePptx` and `PptxDocument` remain unchanged during early V2
work. Introduce an explicitly experimental scene API first, for example:

```ts
parsePptxScene(input, options): Promise<PptxSceneDocument>
parsePptxSceneWithDiagnostics(input, options): Promise<PptxSceneParseResult>
```

The exact names are frozen only after public-contract review. V1 output is then
derived through a tested `toPptxDocumentV1(scene)` adapter where practical.
During migration:

- HTML text is derived from structured paragraphs and runs;
- `layoutElements` is derived from the selected layout and master;
- flattened group coordinates are produced only by the V1 adapter;
- resolved hex colors remain available to V1;
- compatibility spellings such as `shapType` are retained in V1;
- V1 `order` is derived from V2 ordered arrays;
- existing reader tests remain green.

No public API switch occurs until the V2 parser covers the current V1 feature
matrix, the adapter is semantically equivalent for that matrix, and migration
documentation exists.

### Parse authored and resolved views together

Domain parsing should return both:

1. authored values and their owner; and
2. a resolved view for rendering and V1 compatibility.

Resolution remains deterministic, but it must no longer erase provenance. For
example, a shape can report that its text color is `scheme:accent1` inherited
from a layout placeholder while also exposing the resolved RGB preview.

### Stable key assignment

Keys are deterministic for a source snapshot but remain independent of source
relationship IDs. Assignment inputs may include the owner's stable key, source
role, shape identity, and authored order. Duplicate or malformed source IDs are
diagnosed and disambiguated internally.

Key stability requirements:

- the same unedited source produces the same keys across sequential and
  concurrent parses with the same schema version;
- keys remain stable after JSON serialization and validation;
- added objects receive keys derived from operation IDs or caller keys;
- keys never reveal unsafe package paths;
- keys are not reused within one edit sequence after deletion.

### Capability-aware parsing

The scene reader records a bounded support classification for features that
matter to later edits:

```text
understood | preserved-opaque | omitted-with-diagnostic | security-rejected
```

This classification is internal to package planning and summarized in the
round-trip support profile. It does not place arbitrary source XML in the
semantic document.

## Proposed public APIs

Names are provisional, but the separation is required.

### Scene parser and source-free creation

```ts
export async function parsePptxScene(
  input: PptxInput,
  options?: PptxSceneParseOptions,
): Promise<PptxSceneDocument>;

export async function createPptx(
  document: PptxSceneDocument,
  options?: PptxCreateOptions,
): Promise<PptxWriteResult>;

export function validatePptxScene(
  value: unknown,
  options?: PptxValidationOptions,
): PptxValidationResult;
```

`createPptx` accepts no hidden source package. Unknown or preservation-only
elements are rejected. If canonical hierarchy arrays are empty, the writer
creates one documented minimal theme/master/layout hierarchy instead of placing
synthetic duplicates on every slide. Partially linked hierarchy objects are
invalid rather than silently repaired.

### Source-preserving round-trip

```ts
export async function readPptxRoundTrip(
  input: PptxInput,
  options?: PptxRoundTripReadOptions,
): Promise<PptxRoundTripSnapshot>;

export async function applyPptxEdits(
  snapshot: PptxRoundTripSnapshot,
  operations: readonly PptxEditOperation[],
): Promise<PptxRoundTripSnapshot>;

export async function writePptxRoundTrip(
  snapshot: PptxRoundTripSnapshot,
  options?: PptxRoundTripWriteOptions,
): Promise<PptxWriteResult>;

export async function serializePptxRoundTripJson(
  snapshot: PptxRoundTripSnapshot,
): Promise<PptxRoundTripPortableJson>;

export async function parsePptxRoundTripJson(
  value: unknown,
): Promise<PptxRoundTripSnapshot>;
```

The runtime snapshot keeps source bytes without forcing Base64 expansion:

```ts
interface PptxRoundTripSnapshot {
  schemaVersion: 1;
  format: 'pptx';
  source: {
    kind: 'bytes';
    data: Uint8Array | Blob;
    sha256: string;
    byteLength: number;
    conformance: 'transitional' | 'strict' | 'unknown';
  };
  supportProfile: PptxSupportProfile;
  document: PptxSceneDocument;
  operations: PptxEditOperation[];
  consistency: PptxSnapshotConsistency;
}

interface PptxRoundTripPortableJson {
  schemaVersion: 1;
  format: 'pptx';
  source: {
    kind: 'base64';
    packageBase64: string;
    sha256: string;
    byteLength: number;
    conformance: 'transitional' | 'strict' | 'unknown';
  };
  supportProfile: PptxSupportProfile;
  document: PptxSceneDocument;
  operations: PptxEditOperation[];
  consistency: PptxSnapshotConsistency;
}

interface PptxSnapshotConsistency {
  contractVersion: string;
  keyAlgorithmVersion: string;
  canonicalizationVersion: string;
  capabilityProfileVersion: string;
  hashAlgorithm: 'sha256';
  sourceManifestSha256: string;
  semanticPreviewSha256: string;
  operationsSha256: string;
}
```

The portable JSON form exists for storage and transport and has stricter size
limits because Base64 and JavaScript strings amplify memory. Runtime editing
uses bytes or `Blob` directly. Both forms include exact contract, key-assignment,
canonicalization, capability-profile, and hash-algorithm versions.

Snapshot consistency binds the source hash, schema version, support profile,
stable target-key manifest, semantic preview hash, and ordered operation list.
`applyPptxEdits` is asynchronous so browser implementations can use Web Crypto
without a synchronous hashing fallback.

### Snapshot threat model

An unkeyed SHA-256 value detects corruption and stale state; it does not prove
authenticity against a party that can edit JSON and recompute a hash. The public
contract must not describe this as tamper-proof security.

The actual write security boundary is:

1. verify and strictly reparse the source bytes;
2. rebuild stable source targeting and the package/reference graph;
3. validate the versioned consistency binding;
4. treat the operation list, not the semantic preview, as edit authority;
5. replay every operation with preconditions;
6. calculate dirty closure and preservation conflicts from the reparsed source;
7. verify the completed output independently.

Direct mutation of snapshot objects is unsupported and detected as a
consistency failure when possible, but cryptographic authenticity requires a
separate keyed MAC or signature managed by a trusted caller/service. OAKit does
not own that secret.

### Edit operation families

Initial operation families should be narrow and composable:

```text
presentation
  setDocumentProperties
  setSlideSize

slides
  addSlide | duplicateSlide | deleteSlide | reorderSlide
  setSlideLayout | setSlideBackground | setSlideHidden
  setSlideTransition | setSpeakerNotes

elements
  addElement | deleteElement | duplicateElement
  setElementTransform | setElementVisibility
  moveElementBefore | moveElementAfter
  moveElementToGroup | ungroupElement

text
  replaceTextRange | replaceTextBody
  setParagraphProperties | setRunProperties

shape/media
  setShapeGeometry | setFill | setLine | setEffects
  replaceImage | setImageCrop | replaceAudio | replaceVideo
  connectEndpoint | disconnectEndpoint

table/chart/diagram
  setTableCell | resizeTable | mergeCells | unmergeCells
  setChartSeries | setChartData | setChartFormatting
  setDiagramText (only after SmartArt capability proof)

hierarchy
  addLayout | updateLayout | setPlaceholderOverride
  updateMaster | updateTheme
```

Every operation has:

- a unique `operationId`;
- a stable target key;
- an optional `ifMatch` semantic precondition;
- a bounded typed payload;
- a capability-profile row;
- a deterministic inverse where an inverse is well-defined;
- a declared package-impact rule.

Operations apply sequentially and atomically. One invalid operation aborts the
entire write; no partial result is returned.

`ifMatch` is a versioned canonical semantic digest of the exact target fields
named by the operation contract. It is not an arbitrary JavaScript-object hash.
Canonicalization defines property ordering, Unicode handling, absent versus
explicit values, numeric normalization, media hashes, and referenced-key
ordering.

### Normative operation catalogue

Every public operation requires a specification row before implementation:

```text
target and target owner
precondition fields
semantic effect
identity and allocation effect
reference-graph effect
coordinate-space effect
dirty closure
opaque/MCE conflict rule
inverse or explicit non-invertibility
supported conformance and producer evidence
```

The following high-impact semantics are fixed before general editing begins:

| Operation                     | Required explicit policy                                                                                                                            |
| ----------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `setSlideSize`                | Choose `preserve-coordinates`, `scale-content`, or a future named crop/reflow policy; no implicit scaling                                           |
| `setSlideLayout`              | Preserve slide-owned shapes, deterministically rebind compatible placeholders, identify orphaned overrides, and block ambiguous placeholder matches |
| `duplicateSlide`              | Choose share-versus-clone behavior separately for media, charts, embedded workbooks, notes, comments, and opaque companion parts                    |
| `deleteSlide`                 | Update or block sections, custom shows, slide jumps, zooms, comments, notes, and every known or opaque reference to the slide                       |
| `reorderSlide`                | Preserve stable slide identity and update only order-bearing structures; do not rename parts merely to match order                                  |
| `moveElementToGroup`          | Convert transforms into the target child coordinate space and update connector/animation references without changing visual geometry                |
| `ungroupElement`              | Convert child transforms into owner space, preserve child order, and define connector/animation target behavior                                     |
| `deleteElement`               | Block, detach, or cascade each connector, animation, comment, action, or extension reference according to the operation payload                     |
| `updateTheme`                 | Retain scheme-authored values, recompute derived previews, and never materialize theme-bound values as RGB unless explicitly requested              |
| `updateLayout`/`updateMaster` | Preserve authored owner, calculate all descendant semantic effects, and render representative descendants during verification                       |
| chart-data edits              | State whether formulas, caches, embedded workbook cells, external-data relationships, and recalculation metadata are updated or block the operation |

An operation payload must carry every choice that changes semantics. The writer
must not select a convenient default based on producer behavior discovered at
runtime.

### Complete PPTX reference graph

Before any delete, duplicate, move, regroup, layout, master, or theme operation,
the strict reader builds an owner-scoped reference index covering at least:

- presentation slide, master, layout, notes-master, and handout-master lists;
- sections, custom shows, slide names, slide IDs, and hidden-state ownership;
- layout-to-master, slide-to-layout, notes-to-slide, and notes-master links;
- internal hyperlinks, `ppaction` slide jumps, action settings, and transition
  sounds;
- summary/slide zoom and navigation extensions;
- connector start/end shape IDs and connection-site indexes;
- animation/timing targets, media timing, triggers, and conditions;
- legacy and modern comments, people/authors, and threaded references;
- image, fill, audio, video, poster, SVG/fallback, and embedded-font roles;
- charts, styles, colors, external data, embedded workbooks, and workbook
  relationships;
- SmartArt data, drawing, layout, quick-style, and color parts;
- OLE/ActiveX/macro/signature relationships even when preservation-only;
- custom XML, tags, extension lists, and Markup Compatibility branches;
- every internal and external OPC relationship from its actual owner part.

Known references become typed edges. Unknown markup attached to or referencing
a dirty owner becomes a preservation conflict unless a focused patch rule
proves it remains valid. Absence from the semantic preview is never evidence
that no package reference exists.

## JSON-to-OOXML mapping

The mapping is owned by the writer and tested at public output boundaries. The
table below defines the minimum reverse mapping.

| Semantic owner or element        | Primary OOXML output                                                    | Relationships and companion parts                                                 |
| -------------------------------- | ----------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| Document size and slide manifest | `ppt/presentation.xml`                                                  | `ppt/_rels/presentation.xml.rels`, slide identifiers and order                    |
| Package declarations             | `[Content_Types].xml` and `_rels/.rels`                                 | Default/override MIME declarations and root office-document relationship          |
| Core/custom properties           | `docProps/core.xml`, `docProps/app.xml`, optional `docProps/custom.xml` | Root relationships and content-type declarations                                  |
| Theme                            | `ppt/theme/themeN.xml`                                                  | Master relationship to theme; optional theme-owned media                          |
| Master                           | `ppt/slideMasters/slideMasterN.xml`                                     | Master rels to theme, layouts, and master-owned assets                            |
| Layout                           | `ppt/slideLayouts/slideLayoutN.xml`                                     | Layout rels to master and layout-owned assets                                     |
| Slide                            | `ppt/slides/slideN.xml`                                                 | Slide rels to layout, notes, media, charts, diagrams, links, and embedded objects |
| Background                       | `p:cSld/p:bg` or theme background reference                             | Theme/color-map dependencies when authored as references                          |
| Text box                         | `p:sp`, `p:txBody`, DrawingML paragraphs/runs                           | Hyperlink/action relationships where required                                     |
| Preset/custom shape              | `p:sp`, `p:spPr`, preset or custom geometry                             | Media relationship for picture fills; hyperlink/action rels                       |
| Connector                        | `p:cxnSp` with start/end connection references                          | Target shape identifiers inside the owning shape tree                             |
| Image                            | `p:pic` and an image part                                               | Slide/layout/master rels, media content type, optional external link              |
| Group                            | `p:grpSp` with outer and child transforms                               | Child assets remain related from the owning slide/layout/master part              |
| Table                            | `p:graphicFrame` and `a:tbl`                                            | Table style identity and any text hyperlinks                                      |
| Chart                            | `p:graphicFrame` and `ppt/charts/chartN.xml`                            | Chart rels, embedded workbook, style/color parts, external-data policy            |
| SmartArt                         | `p:graphicFrame` and diagram data/drawing/layout/style/color parts      | Owner rels and cross-part diagram relationships                                   |
| Audio/video                      | Media-bearing `p:pic` or supported media frame                          | Media, poster image, media relationships, timing references                       |
| Math                             | Office Math in text or supported fallback picture                       | Fallback image relationship where present                                         |
| Notes                            | `ppt/notesSlides/notesSlideN.xml`                                       | Notes rels, notes master, slide backlink, notes-owned media                       |
| Transition                       | Slide/layout/master transition node at authored owner                   | Optional sound relationship                                                       |
| Animation                        | Slide timing tree                                                       | Stable target shape IDs and media relationships                                   |
| Comments                         | Version-appropriate comment authors and comment parts                   | Presentation/slide relationships as required by the chosen schema                 |

The complete package inventory also classifies presentation properties,
view properties, table styles, notes size, notes masters, handout masters,
comment authors/people, modern comments, embedded fonts, tags, custom shows,
sections, zoom/navigation extensions, custom XML, signatures, macros, ActiveX,
OLE, and producer extension parts. A domain omitted from source-free creation
must still have an explicit `rejected` or `preservation-only` round-trip policy.

### Markup Compatibility and extension preservation

PowerPoint relies on `mc:AlternateContent`, `mc:Choice`, `mc:Fallback`,
`mc:Ignorable`, `extLst`, and producer namespaces for features including SVG,
modern media, Office Math, comments, animations, and SmartArt.

The reader records every compatibility branch and required namespace in the
internal preservation graph. The writer follows these rules:

- source-free creation emits only a versioned set of understood branches and
  fallbacks;
- a copied owner preserves its compatibility markup byte-for-byte at the part
  payload level;
- a local patch preserves namespace declarations and all untouched branches;
- rebuilding a part is allowed only when every required branch on the dirty
  semantic surface is understood;
- an unknown choice/fallback or extension intersecting a dirty subtree blocks
  before serialization;
- the validator checks that every emitted `Requires` prefix, ignorable prefix,
  fallback representation, and extension relationship remains resolvable.

### Text mapping

The writer maps ordered text nodes as follows:

```text
PptxTextBody.body                  -> a:bodyPr and a:lstStyle
PptxParagraph                     -> a:p
PptxParagraph.properties          -> a:pPr
PptxTextRun                       -> a:r/a:rPr/a:t
PptxTextField                     -> a:fld/a:rPr/a:t
PptxTextBreak                     -> a:br/a:rPr
PptxParagraph.endProperties       -> a:endParaRPr
```

HTML is never parsed back into DrawingML. This avoids ambiguous CSS-to-OOXML
conversion and preserves fields, list semantics, and inheritance.

Text serialization additionally defines `xml:space` behavior, Office
`_xHHHH_` escaping, escaping of literal `_xNNNN_`-shaped strings, invalid XML
control-character rejection, tabs, soft breaks, field identifiers, and Unicode
normalization. These rules are verified against literal XML and native producer
fixtures rather than inferred from the production parser.

### Geometry and unit mapping

- points convert through exactly `12,700` EMUs per point;
- degrees convert through exactly `60,000` angle units per degree;
- DrawingML percentages, fixed-point values, color transforms, and crop values
  each use a domain-specific checked serializer;
- conversions reject non-finite numbers, unsafe integers, negative dimensions,
  and values beyond the declared PPTX limits;
- integer rounding policy is deterministic and covered by boundary tests;
- group children map through their preserved local child coordinate space;
- connector references allocate shape IDs only after the complete owner tree is
  known.

### Master, layout, placeholder, and z-order mapping

The writer serializes each owner's ordered `elements` array directly into its
shape tree. It never merges all hierarchy levels into one slide tree.

Placeholder mapping uses each element's canonical placeholder metadata and
source key. The writer allocates source-independent shape IDs while retaining
the placeholder type, index, orientation, size, prompt behavior, and authored
absence used for inheritance. A local slide override remains a slide-owned
shape; a layout or master edit remains at its authored owner. Resolved preview
transforms or styles are never serialized as local overrides.

### Media mapping

Embedded media is decoded only after lexical Base64 and declared-size checks.
The writer:

- verifies the decoded MIME signature where supported;
- selects a safe extension and content type;
- deduplicates identical bytes by verified hash only when relationship behavior
  remains equivalent;
- never fetches an external target;
- writes external targets only with explicit external relationship mode;
- rejects unsupported active content by default.

Source-preserving mode does not deduplicate existing source parts. Creation
mode may deduplicate only complete representation groups whose roles and
producer behavior are equivalent. Font resources additionally pass embedding
rights and obfuscation-policy checks.

### Chart mapping

Chart formulas and caches remain separate. The capability manifest distinguishes
three modes:

1. literal category/value data and caches without an embedded workbook;
2. formulas plus an OAKit-created embedded workbook;
3. preserved source formulas with embedded or external workbook relationships.

Source-free creation may write mode 1 or 2 only for declared chart kinds. Mode
2 depends on the verified XLSX writer or a separately scoped embedded-workbook
serializer with equivalent package, formula, and recalculation gates; duplicate
workbook implementations are not accepted merely for convenience.

A chart edit specifies whether formulas, caches, embedded workbook cells,
external-data relationships, style/color parts, and recalculation metadata are
updated. Unsupported ChartML or opaque extensions block an edit that reaches
their owner part.

### Math, notes, comments, and presentation-owned parts

Office Math creation uses a supported semantic AST to emit OMML. Parsed LaTeX
and fallback pictures are previews and cannot be used as the sole reverse
mapping. Unsupported OMML is preservation-only and blocks semantic math edits.

Notes slides and notes masters are separate authored owners; there is no notes
layout layer. Handout masters, notes size, presentation/default text styles,
view/show properties, legacy comments/authors, and modern
comments/people/threading each receive explicit model and capability rows
before their creation or edit APIs are advertised.

## Proposed source layout

```text
src/formats/pptx/
├── index.ts                         Supported public entry points
├── types.ts                         V1 compatibility reader types
├── scene-types.ts                   V2 authored semantic scene model
├── parser.ts                        Existing V1 orchestration during migration
├── scene-reader.ts                  Authored plus resolved V2 reader
├── validation.ts                    Public JSON validation and budgets
├── writer/
│   ├── create.ts                    Source-free creation orchestration
│   ├── context.ts                   Per-write registries and allocators
│   ├── package-plan.ts              OPC part and relationship graph
│   ├── reference-index.ts           Slide/shape/part reference graph
│   ├── serialize-xml.ts             Namespace-safe XML serialization
│   ├── markup-compatibility.ts      AlternateContent and extension policy
│   ├── presentation.ts              Manifest, size, sections, properties
│   ├── hierarchy.ts                 Theme/master/layout/placeholder output
│   ├── slide.ts                     Slide output and composite rules
│   ├── notes-comments.ts            Notes masters/slides and comment versions
│   ├── text.ts                      DrawingML structured text
│   ├── shape.ts                     Geometry, connector, groups, transforms
│   ├── fill-line-effect.ts          Authoring style serializers
│   ├── media.ts                     Media registry and relationships
│   ├── table.ts                     Table serialization
│   ├── chart.ts                     Chart and workbook serialization
│   ├── diagram.ts                   SmartArt serialization boundaries
│   ├── animation.ts                 Transitions and timing tree
│   └── verify.ts                    Fresh parse and semantic comparison
└── roundtrip/
    ├── types.ts                     Snapshot, operations, report, limits
    ├── read.ts                      Strict snapshot construction
    ├── portable-json.ts             Bounded Base64 import and export
    ├── validate.ts                  Consistency and schema validation
    ├── apply-edits.ts               Async ordered operation application
    ├── operation-contracts.ts       Normative per-operation semantics
    ├── capability.ts                Multidimensional support matrix
    ├── impact-graph.ts              Semantic and package dirty closure
    ├── part-disposition.ts          Copy/patch/rebuild/add/remove/block plan
    ├── patch.ts                     Proven bounded source-part patching
    ├── write.ts                     Preservation orchestration
    └── verify.ts                    Part hashes, graph, semantics, fidelity
```

No PowerPoint-specific serializer belongs in `src/common`. OPC or XML writer
primitives move to `common` only after a second real format consumes the same
contract.

## Creation-writer pipeline

### 1. Validate semantic JSON

- require the exact supported schema version;
- reject unknown discriminants and unsupported preservation-only elements;
- validate unique stable keys and all references;
- validate authored absence separately from explicit zero, false, null, empty,
  and default-valued attributes;
- validate finite numeric values and unit ranges;
- validate hierarchy links, placeholder references, group cycles, connector
  targets, media keys, and chart dependencies;
- reject round-trip `snapshot` media sources in source-free creation;
- enforce output-growth and traversal limits before allocation;
- do not mutate caller-owned objects or byte arrays.

### 2. Resolve the authoring graph

- preserve explicit theme/master/layout/slide ownership;
- generate one documented minimal theme, master, and layout only when the
  source-free document supplies empty hierarchy arrays;
- resolve derived preview values for semantic verification without replacing
  authored references;
- derive placeholder indexes and composite slide previews from canonical
  element metadata without creating a second editable source;
- calculate shape-tree order and all cross-object dependencies.

### 3. Allocate package parts and identifiers

The per-write context owns:

- part names and content types;
- relationship IDs scoped to each owner part;
- presentation slide IDs;
- shape IDs scoped to each shape tree;
- comment and author IDs where supported;
- media hashes and names;
- chart, notes, layout, master, theme, and embedded-workbook part names.

Allocators are deterministic for equal semantic input but never depend on
mutable module state. Concurrent writes cannot share IDs, caches, or counters.

### 4. Serialize strict OOXML

- serialize through typed namespace-aware builders;
- escape all XML and Office text/attribute values with the declared
  `_xHHHH_`/`xml:space` policy;
- emit only declared namespaces and relationship types;
- emit only supported Markup Compatibility choices, fallbacks, and extensions;
- preserve semantic ordering where the schema makes order meaningful;
- emit Transitional or Strict OOXML only when the selected profile supports
  every emitted feature;
- never concatenate relationship targets or package paths manually.

### 5. Build the OPC package

- write root relationships and content-type declarations from the final graph;
- reject duplicate or case-conflicting part names;
- reject unresolved and orphan internal relationships;
- use deterministic ZIP entry order and timestamps for source-free creation;
- enforce final compressed and expanded output budgets.

### 6. Validate the complete output graph

Before returning bytes:

- reopen the ZIP through an independent fresh instance;
- validate required parts, content types, relationship ownership, target
  existence, uniqueness, and reachability;
- validate XML well-formedness and schema-specific structural invariants;
- run Open XML SDK validation on supported Windows CI profiles and retain
  validator findings as evidence, with understood extension exceptions scoped
  by capability row;
- assert no unsafe internal target or undeclared external target exists;
- assert every allocated slide, shape, layout, master, and relationship ID is
  valid in its scope.

### 7. Verify semantic intent

Strictly parse the generated output through a fresh scene-reader instance and
compare its normalized supported semantics with the input. Comparisons account
only for documented canonicalization such as EMU rounding, generated default
hierarchy, and media deduplication.

The production writer and production reader cannot be the only oracle.
Independent OOXML assertions and producer/render evidence are required before a
feature reaches C2.

## Source-preserving round-trip pipeline

### 1. Validate runtime or portable snapshot and budgets

- validate schema and support-profile versions;
- validate Base64 lexically before decoding only for the portable JSON form;
- verify declared byte length and SHA-256;
- enforce JSON depth, object, array, string, media, operation, and decoded
  source limits;
- reject duplicate keys, operation IDs, invalid references, and unknown
  operation fields;
- verify the consistency binding before using the semantic preview;
- enforce separate portable-JSON and runtime-byte memory profiles.

### 2. Strictly reparse the source

The writer never trusts only the serialized preview. It reparses the verified
source package with round-trip strictness, rebuilds stable targeting metadata,
and compares the bound source semantic manifest. A mismatch causes
`snapshot-consistency-failed`.

### 3. Replay operations atomically

- check operation discriminants and `ifMatch` preconditions;
- apply operations in order to immutable semantic state;
- validate local invariants after each operation;
- validate the entire hierarchy and reference graph at the end;
- calculate deterministic inverses where supported;
- abort the sequence on the first typed failure.

### 4. Calculate the dirty dependency closure

The impact graph starts with semantic targets and expands through package
owners and references. Representative rules include:

| Edit                          | Minimum dirty closure                                                                                                                    |
| ----------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| Replace text in a slide shape | Owning slide XML; slide rels only if hyperlinks/actions change                                                                           |
| Change a layout placeholder   | Layout XML plus semantic verification of every dependent slide; dependent slides stay copied if no serialized override is required       |
| Change a master logo          | Master XML, relevant master rels, new/removed media, content types; dependent slides/layouts remain copied when references do not change |
| Change a theme color          | Theme XML; any dirty cached/resolved artifacts required by the supported profile; preserve authored scheme references elsewhere          |
| Change a slide layout         | Slide relationship, placeholder-rebinding plan, and every extension/reference affected by the old or new layout                          |
| Reorder shapes                | Only the owning slide/layout/master shape tree, unless animation or connector references require companion validation                    |
| Move an element into a group  | Owning shape tree and any connector/animation targets whose IDs or coordinate spaces change                                              |
| Replace an image              | Owning rels, media part, content types when MIME changes, and owner XML if crop or geometry changes                                      |
| Add/delete/reorder a slide    | Presentation manifest/rels, slide companions, content types, sections, custom shows, slide jumps, zooms, comments, and opaque references |
| Edit notes                    | Notes slide and its rels; notes master remains copied unless the operation targets it                                                    |
| Edit chart data               | Chart XML, chart rels, caches, embedded workbook and its graph as required by the chart support profile                                  |
| Edit SmartArt                 | Diagram data/drawing/layout/style/color closure; block if an opaque algorithm or extension is reached                                    |
| Change a connector endpoint   | Owning shape tree; validate target shape IDs and animations referencing affected shapes                                                  |

Every opaque extension or Markup Compatibility branch attached to a dirty owner
or referencing a changed identity is a conflict until a focused rule proves it
can survive the selected patch or rebuild unchanged.

### 5. Assign part dispositions

Every source and planned part receives exactly one disposition:

| Disposition | Meaning                                                                                                                               |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| `copy`      | Preserve the original uncompressed part payload exactly because the part is proven unaffected; ZIP record metadata may be regenerated |
| `patch`     | Apply a bounded, typed, position-safe modification that preserves all other source markup in the dirty part                           |
| `rebuild`   | Serialize the complete dirty part because its supported semantic surface is fully understood                                          |
| `add`       | Create a new reachable part and required declarations/relationships                                                                   |
| `remove`    | Remove a part only after proving no remaining relationship or opaque owner depends on it                                              |
| `block`     | Refuse the write because preservation or semantic correctness cannot be proven                                                        |

Local patching is allowed only for an operation-specific XML structure with
unambiguous identity, ownership, cardinality, namespace behavior, and tests.
String replacement over raw XML is forbidden.

### 6. Handle no-op exactly

If the operation list is empty, consistency is valid, and the derived source
manifest is unchanged, return the original decoded bytes directly. Do not
rebuild the ZIP, normalize XML, change timestamps, reorder parts, or rewrite
relationships. This is the R0 guarantee.

### 7. Serialize the edit plan

- copy unaffected uncompressed part payloads exactly;
- patch or rebuild only the proven dirty closure;
- allocate new identifiers without colliding with preserved identifiers;
- update content types and owner-scoped relationships from the final graph;
- remove only proven unreachable parts;
- preserve or correctly regenerate every required Markup Compatibility prefix,
  branch, fallback, and extension relationship;
- retain the source conformance mode unless a separately named, fully verified
  conversion operation exists.

### 8. Validate output and preservation evidence

The writer validates:

- the complete OPC graph and XML structure;
- exact hashes of every uncompressed part payload classified `copy`;
- requested semantic postconditions through a fresh strict parse;
- absence of unintended semantic changes in supported unaffected domains;
- operation-specific package invariants;
- declared producer/render requirements for R3.

### 9. Return bytes and a fidelity report

```ts
interface PptxWriteResult {
  data: Uint8Array;
  report: {
    level: 'C1' | 'C2' | 'R0' | 'R1' | 'R2' | 'R3';
    supportProfile: {
      id: string;
      version: string;
      effectiveLevel: 'C1' | 'C2' | 'R0' | 'R1' | 'R2' | 'R3';
      producerMatrix: string[];
    };
    operations: PptxOperationEvidence[];
    copiedPartCount: number;
    patchedPartCount: number;
    rebuiltPartCount: number;
    addedPartCount: number;
    removedPartCount: number;
    diagnostics: PptxWriteDiagnostic[];
    producerEvidence: string[];
  };
}
```

Part provenance in reports is bounded and sanitized. It does not expose raw
source XML or reusable internal package objects.

## Security and preservation policy

### Default `reject-active` mode

The default creation and edit writers reject or block:

- encrypted, password-protected, IRM/RMS-protected, or otherwise inaccessible
  packages with distinct typed diagnostics;
- VBA projects and macro-enabled output;
- ActiveX controls and executable embedded objects;
- scripts or launch actions;
- unsafe or undeclared external relationship targets;
- digital-signature-preserving claims after a content edit;
- unsupported OLE mutations;
- any operation that would require fetching an external resource;
- any unknown extension that intersects a dirty owner without a preservation
  proof.

Existing external hyperlinks may be preserved in round-trip mode only under an
explicit policy and are never opened or fetched.

Hyperlinks and actions are separate typed domains. The writer allowlists safe
URI protocols and supported internal slide-jump actions. Macro, program, OLE,
file-launch, and unknown `ppaction` schemes are active content and are rejected
or preserved only outside every dirty closure under the selected policy.

### Opt-in opaque preservation

An opt-in `preserve-opaque` policy may copy macros, OLE parts, vendor
extensions, or other opaque parts only when:

- the package is not security-rejected;
- the part and all connecting relationships remain outside the dirty closure;
- no operation changes the owner's semantics or reference identity;
- the fidelity report identifies the preserved opaque feature class;
- no claim is made that OAKit understood or validated its behavior.

If any condition cannot be proved, the operation is blocked.

### Digital signatures

Returning exact source bytes at R0 preserves existing signatures because no
byte changes. Any content edit invalidates a package signature. The default
behavior is to block a signed-package edit with a typed diagnostic unless a
future explicit `removeSignatures` operation is requested. OAKit does not claim
to re-sign packages.

### Macro-enabled presentations

`.pptm` creation and editing are separate future capabilities. The PPTX writer
must not silently preserve or emit a macro project under a `.pptx` claim.

### Encrypted and rights-managed packages

Standard encrypted Office presentations may be compound binary containers
rather than directly readable OPC ZIP packages. The PPTX API does not guess a
password, decrypt content, or bypass rights management. It reports
`encrypted-package` or `protected-package` before semantic parsing. Future
credential-assisted decryption must be a separate caller-authorized boundary
and cannot weaken round-trip limits or active-content policy.

### Untrusted document text

Text, notes, links, comments, alt text, metadata, and embedded filenames are
data, never host instructions. The writer does not execute them, evaluate
fields, make network requests, launch media, or invoke an Office producer.
Producer verification runs only in controlled CI fixtures.

## Resource limits

Writer limits extend the existing parser limits and apply before expensive
allocation where possible:

- maximum JSON bytes, nesting depth, object count, array items, and string
  bytes;
- maximum masters, layouts, slides, elements per owner, group depth, and total
  elements;
- maximum paragraphs, runs, fields, text bytes, and list depth;
- maximum geometry guides, paths, commands, points, and connector references;
- maximum table rows, columns, cells, merges, and rich-text nodes;
- maximum charts, series, points, formulas, caches, and embedded workbook bytes;
- maximum media entries, per-media decoded bytes, and total decoded media;
- maximum operations, payload bytes, dirty parts, and dependency-graph edges;
- maximum generated parts, relationships per owner, XML bytes per part, total
  expanded output, and compressed output bytes;
- maximum verification parse work and diagnostics.

Runtime and portable snapshot limits are separate. Portable Base64 is budgeted
for encoded characters, decoded bytes, and JavaScript string amplification.
Runtime snapshots avoid Base64 but still budget simultaneous source bytes,
expanded ZIP state, semantic graph, dirty parts, verification state, and output
bytes. Implementations must not assume all of these copies fit merely because
the compressed source is under `maxInputBytes`.

Benchmark tiers include at least small (1 MiB), normal (25 MiB), and large
(100 MiB) source packages with producer-realistic media and XML ratios. Each
tier records peak browser/Node memory, parse time, operation-plan time, write
time, and verification time. Advertised limits are no higher than the smallest
environment that passes the declared runtime matrix.

Limit violations are fatal and return no package. Defaults must work in Node.js
and browsers without relying on filesystem-only spill behavior.

## Errors and diagnostics

Writer diagnostics use stable codes. At minimum:

```text
invalid-scene-document
unsupported-schema-version
invalid-snapshot
snapshot-consistency-failed
operation-precondition-failed
invalid-operation-target
unsupported-edit-operation
unsupported-feature
unsupported-conformance
invalid-hierarchy-reference
invalid-element-reference
duplicate-public-key
reference-conflict
invalid-media
invalid-numeric-value
invalid-office-text-escape
resource-limit-exceeded
opaque-preservation-conflict
markup-compatibility-conflict
signed-package-edit-blocked
encrypted-package
protected-package
unsafe-external-relationship
package-graph-invalid
xml-serialization-failed
semantic-verification-failed
fidelity-not-achieved
producer-verification-failed
```

A diagnostic may identify the operation ID, stable public key, slide/layout/
master key, feature class, sanitized owner part, and requested fidelity. It must
not require callers to parse a message string or inspect internal XML.

## Feature and operation support matrix

Reader support does not automatically imply writer or edit support. A single
domain/operation flag is too coarse: the same image edit may be R3 for an
embedded PNG on a slide, R2 for a cropped JPEG on a layout,
preservation-only for an SVG/fallback group on a master, and unsupported for an
unknown linked-media extension.

The versioned capability manifest records these dimensions:

```text
domain
× feature subtype
× operation
× authored owner
× source conformance
× preservation/security mode
× producer and version
→ capability and evidence
```

Each resulting row is classified as:

```text
unsupported | preservation-only | create-C1 | create-C2
edited-R1 | edited-R2 | edited-R3
```

The effective fidelity of one write is the minimum proven level across every
feature subtype, operation, owner, conformance mode, opaque dependency, and
producer claim used by that output. A high-confidence basic edit cannot hide an
unsupported companion feature in the same dirty closure.

The planned progression is:

| Domain                         | Creation target                                               | First edited operations                          | Evidence required for highest declared level                             |
| ------------------------------ | ------------------------------------------------------------- | ------------------------------------------------ | ------------------------------------------------------------------------ |
| Document/slide hierarchy       | Size, theme, one or more masters/layouts, ordered slides      | Add/delete/reorder/duplicate slide, set layout   | Manifest/relationship graph, placeholder inheritance, producer open/save |
| Text/placeholders              | Structured paragraphs, runs, fields, bullets, links, fitting  | Replace text/range, set paragraph/run properties | DrawingML assertions, V1 adapter parity, rendering and font corpus       |
| Basic shapes                   | Preset geometry, fills, lines, effects                        | Transform, reorder, restyle, replace geometry    | EMU/angle boundaries, geometry corpus, visual diffs                      |
| Custom geometry/connectors     | Guides, paths, endpoints                                      | Adjust geometry and connector endpoints          | Connection-site and group-coordinate properties, producer rendering      |
| Groups                         | Nested local coordinate spaces                                | Group, ungroup, reorder/move children            | Nested transform properties and mutation tests                           |
| Images                         | Embedded/external policy, crop, effects                       | Add/replace/remove/crop                          | MIME/signature, rel/media/content-type graph, render diffs               |
| Media representations/fonts    | Declared fallback groups and permitted embedded fonts         | Replace complete supported group                 | MCE/fallback roles, font rights/obfuscation, producer rendering          |
| Tables                         | Cells, merges, sizes, styles                                  | Edit cell, resize, merge/unmerge                 | Independent table XML checks and producer corpus                         |
| Charts                         | Declared chart kinds, formulas/caches, optional workbook      | Set supported data and formatting                | Chart/workbook graph, cache policy, producer/render matrix               |
| Audio/video                    | Embedded media and poster image for declared variants         | Add/replace/remove, timing after support         | Media/timing relationships and producer playback metadata                |
| Notes                          | Structured speaker notes                                      | Replace notes                                    | Notes-master/notes-slide relationships and producer reopen               |
| Transitions                    | Declared transition types and duration                        | Set/remove transition                            | XML and producer behavior corpus                                         |
| Animations                     | Preservation initially                                        | Narrow target/timing operations later            | Full timing-tree reference and producer matrix                           |
| SmartArt                       | Preservation first; creation only for proven layouts          | Text edits only after dependency proof           | All diagram parts, visual/producer evidence, opaque conflict tests       |
| Math                           | Semantic AST to supported OMML; LaTeX remains derived         | Replace supported expression                     | OMML assertions, AST fidelity, fallback-image and opaque conflict policy |
| Comments/sections/custom shows | Add only after explicit model                                 | Narrow owner-specific operations                 | Cross-part IDs, references, producer compatibility                       |
| Accessibility metadata         | Title, description, decorative state, language, reading order | Set supported accessibility fields               | Literal nonvisual properties, assistive metadata, producer reopen        |
| Presentation/view/handout data | Declared presentation defaults and supported owner parts      | Narrow property operations                       | Part ownership, notes size, view/show behavior, producer compatibility   |
| OLE/macros/unknown extensions  | Rejected or preservation-only                                 | None initially                                   | No edit claim until dedicated security and producer design exists        |

The shipped capability is the intersection of implemented rows, passing
evidence, and documentation. A partially implemented advanced domain never
lowers the strictness of basic domains.

## Independent test strategy

### Test independence rules

- Tests construct minimal OOXML fixtures using literal XML or a test-only
  package builder, not production writer helpers.
- Writer assertions inspect literal expected parts, attributes, relationships,
  order, and media hashes before reparsing.
- Parser/writer round-trip is a property, not the only oracle.
- Windows CI runs Open XML SDK validation for supported conformance profiles;
  XML/OPC validators and native producers remain independent additional oracles.
- Expected semantic values are written independently rather than copied from
  production types at runtime.
- Every discovered fuzz or mutation counterexample becomes a minimized fixed
  regression fixture.

### V2 parser parity and provenance

For every current reader domain:

- assert authored owner and resolved preview independently;
- assert authored absence remains absent after parsing and source-preserving
  edits, even when the resolved preview contains a value;
- assert placeholder linkage across master, layout, and slide;
- assert placeholder indexes are derived from canonical element metadata and
  cannot conflict with a second public definition;
- assert composite rendering order and placeholder substitution across
  background, master, layout, and slide layers;
- assert canonical z-order in every owner and nested group;
- assert structured text nodes and the derived V1 HTML;
- assert local and flattened group transforms separately;
- assert authored color source and resolved hex preview;
- assert stable keys across repeated and concurrent parsing;
- assert V1 adapter output against the existing public contract.

### Source-free creation properties

The primary semantic property is:

```text
normalize(parsePptxScene(createPptx(document)))
  == normalize(document)
```

This is asserted only after independent package checks and only for fields in
the supported creation profile. Additional properties include:

- equal input yields deterministic package semantics and deterministic bytes;
- equivalent point values round to the documented EMU representation;
- ordered arrays produce the same shape-tree order;
- inserting one element changes only the intended order and allocations;
- adding an unused media resource is either rejected or deterministically
  omitted according to the contract;
- concurrent writes never share counters or cache state;
- invalid input produces no partial bytes.

Creation fixtures additionally inspect required shape-tree root nodes,
owner-scoped nonvisual IDs, placeholder attributes, Markup Compatibility
branches, Office text escaping, and source-independent hierarchy allocation.

### R0 exactness

For every accepted source corpus entry:

```text
source bytes
  -> readPptxRoundTrip
  -> serializePptxRoundTripJson
  -> ordinary JSON stringify/parse
  -> parsePptxRoundTripJson
  -> writePptxRoundTrip(no operations)
  -> bytes exactly equal source bytes
```

Compare byte length and SHA-256 first, then full byte equality. Cover ZIP entry
orders, timestamps, compression variants, Strict and Transitional packages,
unknown extensions, and opt-in opaque content under the declared policy.
Run the same no-op property directly through the runtime byte/Blob snapshot so
portable serialization and runtime preservation are independently verified.

### Part preservation after edits

Each edit fixture defines literal expected:

- operation and precondition;
- semantic target state;
- dirty dependency closure;
- part disposition for every affected source part;
- exact hashes for all copied uncompressed part payloads;
- required graph and XML changes;
- fidelity level or typed blocking diagnostic.

No test may accept rebuilding an extra part merely because the output reparses.

### Producer corpus

Maintain versioned fixtures produced by:

- current supported Microsoft PowerPoint on Windows;
- current supported Microsoft PowerPoint on macOS;
- LibreOffice Impress;
- Google Slides export;
- selected historical producer versions when a regression justifies them.

The corpus must include native authoring rather than only synthetic XML:

- multiple themes, masters, layouts, and placeholder chains;
- reordered, hidden, sectioned, and differently sized presentations;
- mixed fonts, scripts, bidirectional text, fields, lists, tabs, and autofit;
- nested groups, rotations, flips, custom geometry, and connectors;
- images with crop/effects, SVG/EMF fallbacks where supported;
- complex tables and chart variants with embedded workbook data;
- SmartArt, equations, notes, transitions, animations, audio, and video;
- comments, hyperlinks, external relationships, custom XML, and extensions;
- large but valid boundary documents.

R3 evidence requires open without repair, controlled save/reopen, semantic
inspection, and retained producer/version artifacts. Google Slides verification
may use controlled import/export fixtures rather than network access inside
core tests.

On supported Windows CI, retain Open XML SDK validation output and PowerPoint
repair-log evidence. LibreOffice headless validation and Google Slides
import/export have producer-specific expected normalization; they are not used
as a reason to require structurally identical post-save packages.

### Visual verification

For C2 and visual R3 rows:

- render source and output with the selected producer or controlled renderer;
- compare source and edited output with the same renderer/environment before
  making cross-producer comparisons;
- compare slide dimensions and object bounding boxes;
- use pixel or perceptual diffs with feature-specific thresholds;
- normalize environmental variables such as installed fonts and rendering DPI;
- retain diff images for failures;
- pair every visual assertion with semantic/package assertions so a tolerant
  screenshot threshold cannot hide structural corruption.
- after master, layout, theme, placeholder, embedded-font, or shared-style
  edits, render representative descendant slides for every affected branch.

### Property and fuzz testing

Use generators for:

- valid and invalid scene JSON, discriminants, references, and hierarchy cycles;
- point/EMU/angle/percentage numeric boundaries and non-finite values;
- slide/layout/master trees and placeholder inheritance;
- nested groups, coordinate spaces, z-order operations, and connector graphs;
- text runs, fields, Unicode, bidirectional text, XML escaping, bullets, tabs,
  and hyperlinks;
- significant whitespace, literal `_xNNNN_` text, invalid XML controls, field
  IDs, and `xml:space` boundaries;
- custom paths, guides, handles, stops, effects, and line variants;
- tables, merges, charts, series, formulas, and caches;
- media MIME mismatches, malformed Base64, duplicate hashes, and size limits;
- relationship graphs, missing targets, duplicate IDs, cycles, and unsafe paths;
- ordered operation sequences, stale `ifMatch`, duplicate IDs, inverses, and
  non-overlapping commutativity;
- unknown extensions attached inside and outside dirty closures;
- `mc:AlternateContent`, choice/fallback, ignorable-prefix, SVG/fallback, and
  extension-relationship combinations;
- encrypted, password-protected, rights-managed, signed, macro-enabled, and
  malformed container boundaries;
- ZIP/XML bombs and output-amplification attempts.

Metamorphic properties include:

- empty operations always select R0 and preserve bytes exactly;
- applying an operation and its supported inverse restores semantic state;
- independent operations commute semantically even when package identifiers
  differ;
- reordering twice to the original position restores canonical order;
- group then supported ungroup preserves visual transforms within rounding;
- theme changes affect scheme-bound previews but not authored RGB colors;
- repeated validation and write calls do not mutate caller data;
- sequential and concurrent calls are observably deterministic.

### Mutation testing

Mutation testing must cover all new production writer and round-trip modules,
not only parser helpers. The required mutation score, covered mutation score,
and test-strength score are all 100% for the declared scope: zero survived,
zero no-coverage, and zero timed-out mutations.

High-risk mutation targets include:

- unit conversions and rounding boundaries;
- ordered-array and z-order comparisons;
- placeholder precedence and owner selection;
- authored-absence preservation and composite hierarchy ordering;
- group coordinate transforms;
- color-transform order and defaults;
- XML escaping and namespace emission;
- relationship target ownership and external-mode checks;
- ID allocation, collision handling, and stable key validation;
- operation ordering and `ifMatch` preconditions;
- dirty-closure graph edges and part dispositions;
- opaque conflict, signature, active-content, and limit guards;
- Markup Compatibility, Office text escaping, encrypted/protected-package, and
  media-representation policies;
- copied-part hash verification;
- fidelity-level classification;
- final graph and semantic verification.

Timeouts fail the gate and are investigated until killed or removed through a
proved equivalent-mutant exclusion. Equivalent mutations require a documented
proof and a narrowly scoped configuration exception only when the expression
is truly semantically identical.

### Browser and package verification

- Run parser, creation writer, snapshot JSON serialization, edit writer, and
  verification in real Chromium.
- Verify there is no Node-only import in the browser PPTX writer bundle.
- Test Blob, ArrayBuffer, and Uint8Array boundaries without caller mutation.
- Run ESM, CommonJS, root export, and `@evoelsewhere/oakit/pptx` package tests.
- Run the Node 20/22/24 matrix and current browser matrix.
- Enforce bundle and memory budgets separately from correctness tests.
- record peak memory and time for the 1 MiB, 25 MiB, and 100 MiB benchmark tiers
  for runtime snapshots and the permitted portable-JSON tiers.

## Delivery sequence

Each phase starts with independent failing evidence and ends with focused,
regression, type, lint, browser, fuzz, corpus, and mutation gates proportional
to the change. The existing PPTX reader remains releasable throughout.

### 1. Freeze terminology and support reporting

- document C1/C2 and R0-R3;
- define the multidimensional capability-manifest schema;
- define unsupported versus preservation-only behavior;
- add public-contract tests that forbid unearned capability claims.

Gate: documentation review and contract tests.

### 2. Freeze authored, resolved, placeholder, and composition contracts

- define authored absence versus explicit values;
- define resolved previews and bounded value-source metadata;
- make element placeholder metadata canonical and parser indexes derived;
- define background/master/layout/slide composition and placeholder
  substitution;
- define descendant verification after hierarchy edits.

Gate: model fixtures, placeholder/inheritance properties, render-composition
tests, and mutation targets.

### 3. Freeze snapshot transport and threat model

- separate runtime byte/Blob snapshots from portable Base64 JSON;
- version source hashes, key assignment, canonicalization, and support profiles;
- define consistency checks without claiming unkeyed-hash authenticity;
- make edit application asynchronous where Web Crypto is required;
- define memory amplification and transport limits.

Gate: browser/Node serialization, corruption, stale-state, memory, concurrency,
and cryptographic-boundary tests.

### 4. Freeze operation semantics and the complete reference graph

- write normative contracts for every initial operation;
- define `ifMatch` canonical target digests;
- define slide-size, layout-rebind, duplicate, delete, group/ungroup, theme, and
  chart policies;
- index all slide, shape, timing, comment, media, chart, SmartArt, extension,
  and owner-part references;
- define Markup Compatibility and opaque-conflict behavior.

Gate: operation fixtures, literal dirty closures, graph fuzzing, MCE conflicts,
inverse properties, and mutation tests.

### 5. Freeze the complete V2 scene invariants

- add model-only test fixtures independent of parser code;
- define hierarchy, stable keys, ordered arrays, local coordinate spaces,
  structured text, authored styles, representation-group media, comments,
  notes masters, presentation-owned parts, and limits;
- define math AST and advanced-domain preservation boundaries;
- add strict semantic JSON validation tests before implementation.

Gate: types, schema/validator tests, fuzzed invalid JSON, mutation scope.

### 6. Build authored master/layout/slide parsing

- parse explicit themes, masters, layouts, and slides;
- retain authored absence, backgrounds, color maps, flags, and canonical
  placeholder metadata;
- provide resolved views without erasing provenance;
- prove deterministic stable keys.

Gate: native producer fixtures, inheritance/composition properties, descendant
rendering, and concurrency tests.

### 7. Replace HTML-only text normalization with structured text

- parse paragraphs, runs, fields, breaks, bullets, tabs, spacing, links, and
  text-body properties;
- preserve significant whitespace, Office escapes, field IDs, and invalid
  control-character diagnostics;
- derive V1 HTML through an adapter;
- preserve current V1 text semantics.

Gate: independent DrawingML text fixtures, V1 parity, Unicode/browser/fuzz/
mutation tests.

### 8. Preserve geometry and coordinate spaces

- add authored transforms and group child spaces;
- keep SVG paths and flattened V1 coordinates as derived views;
- distinguish connectors and preserve endpoint references;
- retain custom-geometry authoring inputs.

Gate: nested-group and connector corpus, numeric properties, visual checks.

### 9. Preserve authored style, media groups, and compatibility markup

- implement theme-aware colors and ordered transforms;
- add authoring fill, line, and effect unions;
- introduce representation-group media with SVG/raster and poster/payload roles;
- add embedded-font rights/obfuscation policy;
- preserve and classify `AlternateContent`, `Ignorable`, and extension lists;
- add encrypted/protected package diagnostics;
- keep existing V1 fields through adapters.

Gate: style/media/fallback corpus, MIME/MCE/security fuzz, browser and mutation
tests.

### 10. Enrich tables, charts, diagrams, math, notes, comments, and timing

- add structured table text and authored styles;
- preserve chart formulas/caches/workbook ownership and declare chart data mode;
- define SmartArt understood/preservation boundaries;
- add semantic math AST versus OMML preservation boundaries;
- model notes slides, notes masters, legacy/modern comments, and separate
  transition/timing trees.

Gate: one domain at a time with native corpus and no capability overclaim.

### 11. Complete V1 adapter parity

- generate current `PptxDocument` from the V2 scene;
- compare every existing public fixture and diagnostic;
- document deliberate differences only when separately approved;
- decide the long-term public API migration after evidence is complete.

Gate: all current reader, black-box, browser, corpus, fuzz, and mutation suites.

### 12. Build the strict creation package core

- add namespace-safe XML serialization;
- add Office-text escaping, part/content-type/reference graph, MCE handling, and
  scoped allocators;
- create deterministic minimal package/theme/master/layout output;
- add output limits, Open XML SDK validation, and package graph verification.

Gate: independent package tests, XML adversarial suite, browser and mutation.

### 13. Implement basic creation domains

- slides, hierarchy, placeholders, backgrounds;
- structured text;
- basic shapes, connectors, groups, fills, lines, and effects;
- embedded images and hyperlinks.

Gate: parse/create semantic property, literal XML assertions, render and
PowerPoint/LibreOffice/Google Slides producer corpus.

### 14. Implement advanced creation domains

- tables, charts, notes, transitions, math;
- audio/video and poster frames;
- diagrams and animations only row by row after dedicated proof.

Gate: domain-specific package graphs, rendering, producer and mutation suites.

### 15. Implement runtime/portable snapshots and R0

- add runtime byte/Blob source, portable Base64 source, versioned consistency
  binding, support profile, and stable key manifest;
- validate runtime and JSON stringify/parse boundaries independently;
- return exact source bytes for empty operations.

Gate: complete accepted corpus byte equality, adversarial snapshots, browser,
concurrency, memory tiers, fuzz, and mutation.

### 16. Implement operation planning and dirty closure

- add typed normative operations, canonical `ifMatch`, atomic replay, and
  inverses;
- build complete semantic/package/reference impact graph;
- classify copy/patch/rebuild/add/remove/block;
- implement MCE, opaque, encrypted/protected, signature, macro, and
  active-content policies.

Gate: literal closure fixtures, graph fuzzing, security and mutation suites.

### 17. Implement basic edited domains

- text and notes replacement;
- element transforms and z-order;
- basic fill/line changes;
- image replacement;
- slide add/delete/reorder/duplicate and layout selection.

Gate: copied-part hashes, R1/R2 semantic verification, producer R3 evidence per
operation row.

### 18. Implement hierarchy and advanced edited domains

- placeholder/layout/master/theme edits;
- groups and connectors;
- tables and charts;
- transitions and media;
- SmartArt, animations, comments, and extensions only after explicit support.

Gate: complete dirty closures, preservation conflicts, render/producer matrix,
fuzz and mutation.

### 19. Complete producer, validator, and performance evidence

- run Open XML SDK validation and PowerPoint repair-log checks on Windows;
- run PowerPoint Windows/macOS, LibreOffice, and Google Slides
  producer-specific matrices;
- compare source/output with the same renderer and render descendants after
  shared-owner edits;
- enforce small/normal/large runtime and portable memory budgets;
- retain capability/version/evidence artifacts for release classification.

Gate: C2/R3 evidence, zero unexplained repair, declared visual thresholds,
Node/browser budgets, and complete critical mutation scope.

### 20. Integrate packaging, CLI, CI, and documentation

- expose stable subpath APIs without pulling write code into reader-only bundles
  unnecessarily;
- add create and apply commands only after programmatic APIs are stable;
- add Node/browser/corpus/producer/mutation/fidelity CI artifacts;
- update README capability claims only to shipped matrix rows;
- add migration and troubleshooting documentation.

Gate: package compatibility, clean install, CLI black-box, release checklist.

## Recommended atomic commits

The sequence should naturally exceed thirty small commits. Each behavior pair
starts with evidence, then implementation. Normal commit subjects describe the
actual contract; they do not use priority labels or phase names.

```text
docs: define PowerPoint write fidelity
docs: define PowerPoint operation semantics
docs: define PowerPoint snapshot threat model
test: define PowerPoint scene invariants
feat: add PowerPoint scene types
test: reject invalid PowerPoint scene data
feat: validate PowerPoint scene data
test: preserve absent PowerPoint properties
feat: model authored and resolved PowerPoint values
test: preserve PowerPoint hierarchy ownership
feat: parse PowerPoint hierarchy ownership
test: preserve PowerPoint placeholder links
feat: parse PowerPoint placeholder links
test: compose inherited PowerPoint slide layers
feat: compose inherited PowerPoint slide layers
test: preserve structured PowerPoint text
feat: parse structured PowerPoint text
test: preserve PowerPoint text fields
feat: parse PowerPoint text fields
test: preserve PowerPoint Office text escapes
feat: serialize PowerPoint Office text escapes
test: derive compatible PowerPoint text HTML
refactor: derive PowerPoint text HTML from scene data
test: preserve PowerPoint group coordinate spaces
feat: parse PowerPoint group coordinate spaces
test: preserve PowerPoint connector endpoints
feat: parse PowerPoint connector endpoints
test: preserve authored PowerPoint colors
feat: parse authored PowerPoint colors
test: preserve PowerPoint custom geometry
feat: parse PowerPoint custom geometry
test: define PowerPoint media representations
feat: add PowerPoint media representations
test: preserve PowerPoint compatibility branches
feat: index PowerPoint compatibility branches
test: preserve PowerPoint chart sources
feat: parse PowerPoint chart sources
test: preserve PowerPoint notes ownership
feat: parse PowerPoint notes ownership
test: preserve the PowerPoint V1 contract
refactor: derive PowerPoint V1 output from scene data
test: define PowerPoint package writer invariants
feat: add PowerPoint package graph
test: index PowerPoint reference dependencies
feat: add PowerPoint reference index
test: define PowerPoint identifier allocation
feat: allocate scoped PowerPoint identifiers
test: serialize minimal PowerPoint packages
feat: write minimal PowerPoint packages
test: serialize PowerPoint hierarchy
feat: write PowerPoint hierarchy
test: serialize structured PowerPoint text
feat: write structured PowerPoint text
test: serialize PowerPoint shapes and groups
feat: write PowerPoint shapes and groups
test: serialize PowerPoint images
feat: write PowerPoint images
test: verify generated PowerPoint semantics
feat: verify generated PowerPoint semantics
test: validate PowerPoint round-trip snapshots
feat: validate PowerPoint round-trip snapshots
test: serialize portable PowerPoint snapshots
feat: serialize portable PowerPoint snapshots
test: preserve exact PowerPoint source bytes
feat: preserve exact PowerPoint source bytes
test: apply ordered PowerPoint edits
feat: apply ordered PowerPoint edits
test: enforce PowerPoint operation semantics
feat: enforce PowerPoint operation semantics
test: calculate PowerPoint dirty parts
feat: calculate PowerPoint dirty parts
test: block unsafe PowerPoint preservation
feat: enforce PowerPoint preservation policy
test: reject encrypted PowerPoint packages
feat: diagnose encrypted PowerPoint packages
test: edit PowerPoint text and transforms
feat: edit PowerPoint text and transforms
test: edit PowerPoint slides and images
feat: edit PowerPoint slides and images
test: verify copied PowerPoint parts
feat: verify copied PowerPoint parts
test: verify PowerPoint creation in browsers
test: verify PowerPoint edits in browsers
test: verify PowerPoint producer corpus
test: validate PowerPoint with Open XML SDK
test: enforce PowerPoint writer memory budgets
test: enforce PowerPoint writer mutations
ci: verify PowerPoint write fidelity
docs: document supported PowerPoint writing
```

Do not force these exact commits if a change cannot stand alone, but keep tests,
implementation, packaging, CI, and capability documentation reviewable. Never
stage or include unrelated work.

## Quality gates for every advertised domain

An operation/domain row advances only when all applicable gates pass:

1. independent valid, invalid, missing, boundary, security, and concurrency
   tests;
2. literal package graph and XML assertions;
3. fresh strict output reparse and semantic postcondition checks;
4. copied uncompressed-part-payload hash evidence for round-trip edits;
5. property/fuzz tests with minimized regressions;
6. 100% in-scope mutation result with no survivors or untested mutations;
7. real Chromium tests;
8. Node 20/22/24 package matrix;
9. native PowerPoint, LibreOffice Impress, and Google Slides corpus evidence for
   R3/C2 claims;
10. Open XML SDK validation and producer repair-log evidence for applicable
    profiles;
11. same-renderer visual verification and descendant rendering for shared-owner
    edits;
12. authored-absence, MCE/fallback, and copied-part preservation evidence;
13. runtime/portable resource, performance, memory, and output-amplification
    limits;
14. README, API declarations, capability manifest, and implementation in
    agreement.

Fast pull-request CI may run focused subsets. The merge/release gate runs the
complete writer mutation scope, producer corpus appropriate to the claimed
level, package matrix, and browser verification.

## Definition of done

### The parser update is complete when

- V2 preserves authored theme/master/layout/slide ownership;
- V2 preserves authored absence separately from explicit and resolved values;
- every owner and group has one canonical ordered element array;
- placeholder metadata has one canonical public source and composition does not
  duplicate inherited placeholders;
- composite background/master/layout/slide rendering is deterministic;
- elements retain local coordinate spaces and connector references;
- text is structured and HTML is derived;
- colors, geometry, media, charts, tables, notes, and supported advanced domains
  preserve the authoring data required by their declared writer rows;
- stable keys are deterministic and source IDs remain internal;
- the complete package/reference graph classifies known and opaque edges,
  Markup Compatibility branches, and advanced-domain preservation boundaries;
- the V1 adapter preserves the existing public reader contract;
- unsupported versus preservation-only content is observable through bounded
  support evidence and diagnostics;
- all reader regression, browser, corpus, fuzz, concurrency, and mutation gates
  pass.

### Source-free creation is complete at C1 when

- every field in the declared creation schema has a tested OOXML mapping or is
  rejected before serialization;
- the output OPC graph and all XML parts pass strict independent validation;
- Office text escaping, MCE branches/fallbacks, and owner-scoped shape-tree IDs
  pass literal and independent validation;
- a fresh scene parse equals the normalized semantic input;
- generated IDs, relationships, media, and content types are complete and
  collision-free;
- equal input is deterministic and concurrent writes are isolated;
- invalid or over-limit input returns no partial package;
- browser and Node package gates pass.

### Source-free creation is complete at C2 when

- every feature used by the document has passing visual evidence where
  applicable;
- the declared PowerPoint/LibreOffice/Google Slides matrix opens without repair;
- documents using shared masters, layouts, themes, placeholders, or fonts render
  representative descendant slides correctly;
- controlled save/reopen preserves required semantics;
- producer versions and artifacts are retained with the release evidence.

### R0 exact no-op is complete when

- runtime byte/Blob and portable JSON snapshot paths are independently tested;
- portable JSON survives stringify/parse without losing source bytes or keys;
- consistency checks reject corrupted, stale, or directly modified bound state;
- documentation states that unkeyed hashes do not authenticate a maliciously
  recomputed snapshot;
- empty operations return bytes exactly equal to every accepted corpus source;
- the writer does not rebuild ZIP entries, XML, IDs, timestamps, or order;
- opaque and signed packages follow the declared policy;
- browser, concurrency, limits, fuzz, and mutation gates pass.

### An edited domain is complete at R1/R2 when

- every advertised operation has valid, invalid, missing, boundary,
  precondition, conflict, and resource tests;
- the dirty closure is independently asserted;
- every known reference edge is updated or preserved and every unknown/MCE edge
  intersecting the edit is proven safe or blocks;
- every copied uncompressed part payload retains its exact hash;
- every patched or rebuilt part is completely accounted for;
- opaque/MCE conflicts, encryption/protection, signatures, active content, and
  unsafe relationships block correctly;
- output graph validation passes;
- a fresh strict parse proves the requested postcondition and unaffected
  supported semantics;
- no in-scope mutation survives.

### An edited domain is complete at R3 when

- R1 and R2 are complete;
- every advertised operation passes the declared producer matrix;
- the file opens without repair dialogs or repair logs;
- controlled save/reopen does not reveal invalid required structures;
- visual behavior is within the declared thresholds;
- producer/version evidence is retained as a CI or release artifact.

### General PPTX writing may be advertised only when

- capability claims identify creation versus source-preserving edits;
- capability claims are selected from the full subtype/operation/owner/
  conformance/security/producer matrix;
- documentation distinguishes exact no-op from edited fidelity;
- unsupported domains are rejected or explicitly preservation-only;
- there is no best-effort path labeled round-trip;
- implementation, types, tests, support manifest, and README agree;
- the shipped domain/operation matrix has passed all gates required by its
  advertised levels.

## Remaining impossibility boundaries

Even after this plan is complete, OAKit must not promise:

- byte-identical reconstruction from semantic JSON without the opaque source;
- byte-identical ZIP/container output after a content edit, even when every
  unaffected uncompressed part payload is copied exactly;
- cryptographic authenticity from an unkeyed source or snapshot hash;
- safe semantic editing of arbitrary future extension markup;
- retention of a digital signature after changing signed content;
- execution, validation, or re-signing of macros and embedded programs;
- decryption or rights-management bypass without a separate authorized
  credential boundary;
- identical layout when required fonts or producer rendering engines differ;
- identical animation, SmartArt, media, and chart behavior outside the declared
  capability and producer matrix;
- recovery of information that the source producer never stored.

Blocking an unproven edit is a correctness result, not an incomplete write.

## Kickoff exit criteria

The implementation starts only when the preceding reader goal:

- is merged or otherwise frozen at a known commit;
- passes formatting, lint, typecheck, focused and full unit tests;
- passes the applicable browser, corpus, fuzz, package, and mutation gates;
- has no unexplained surviving or untested mutations in its declared scope;
- leaves the worktree free of unrelated changes intended for that goal;
- records any known parser fidelity gaps that this plan must absorb.

These criteria were satisfied by the reader baseline recorded in the kickoff
audit. At each implementation milestone, re-audit the current types, parser
behavior, corpus, and mutation report. This plan is the target architecture;
implementation decisions must be updated when new evidence changes a current
assumption.
