# XLSX Implementation Audit

This audit maps every required outcome from
[`xlsx-reader-plan.md`](xlsx-reader-plan.md) and
[`xlsx-roundtrip-plan.md`](xlsx-roundtrip-plan.md) to implementation and
verification evidence. A row is complete only when the named public contract,
implementation, and independent gate all exist. A passing narrower test does
not complete a broader row.

## Baseline

Audit baseline: branch `feat/xlsx-agent-ready` at `34effbe`, before XLSX code.

| Evidence source        | Baseline finding                                                                                                                                                                                                                                                                               |
| ---------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/formats`          | Only `pptx` exists; there is no XLSX format domain.                                                                                                                                                                                                                                            |
| Root/package/build/CLI | No XLSX export, subpath, build entry, or CLI dispatch exists.                                                                                                                                                                                                                                  |
| Tests and corpus       | No XLSX fixture builder, semantic tests, browser tests, producer corpus, fuzz suite, or mutation targets exist.                                                                                                                                                                                |
| Common package support | Bounded ZIP entry reads, Base64, whole-tree XML, numeric helpers, media typing, and basic OPC target resolution exist. Canonical package identity, content-type parsing, owner-scoped relationship parsing, aggregate archive accounting, and streaming XML are not complete shared contracts. |
| Round-trip support     | PPTX has reusable design evidence for source ownership, canonical hashing, bounded data trees, validation, and exact bytes. There is no XLSX snapshot, operation planner, package graph, writer, or fidelity report.                                                                           |

The two XLSX plans are therefore the authoritative contract. Existing PPTX
behavior is regression evidence and a source of format-neutral primitives, not
evidence that an XLSX row is implemented.

## Reader contract audit

### Public boundary and core behavior

| Requirement                                         | Required evidence                                                                                                   | Status  |
| --------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- | ------- |
| `parseXlsx` and `parseXlsxWithDiagnostics`          | Public black-box calls from `src/formats/xlsx`; root/subpath integration in the final integration commit            | Missing |
| `ArrayBuffer`, `Uint8Array`, and `Blob` equivalence | Literal output equality, caller-byte immutability, Node and Chromium                                                | Missing |
| Option defaults                                     | Contract tests for tolerant errors, supported display text, no image bytes, pivot metadata, and all-sheet selection | Missing |
| Sparse worksheet model                              | Stale/full-grid dimensions never allocate dense state; authored blank styled cells remain distinct                  | Missing |
| Discriminated cell content                          | Exactly one of blank, literal value, or formula plus cached/missing state                                           | Missing |
| Deterministic JSON-compatible output                | No `Date`, non-finite number, raw XML, ZIP object, relationship ID, cache, or process-local object escapes          | Missing |
| Ordered sheet model                                 | Workbook manifest order, worksheet/chart-sheet union, visibility, selected payload state                            | Missing |
| Stable diagnostics                                  | Typed XLSX codes and structured bounded fields; strict/tolerant recovery matrix                                     | Missing |
| Parse isolation                                     | Sequential/concurrent determinism and no process-global mutable state                                               | Missing |

### Package, XML, and selection

| Requirement                 | Required evidence                                                                                                                            | Status  |
| --------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- | ------- |
| Canonical OPC part identity | Traversal, encoded separators/dot segments, invalid percent encoding, query/fragment ambiguity, case sensitivity, duplicate normalized names | Missing |
| Workbook discovery          | Content types plus exactly one internal office-document relationship; relocated Strict and Transitional workbook fixtures                    | Missing |
| Incompatible inputs         | Reject XLSM, XLSB, XLS, encrypted, binary, ambiguous, or missing main parts in both modes                                                    | Missing |
| Relationship ownership      | Owner-relative worksheet, table, drawing, chart, pivot, comment, and external targets; duplicate IDs and missing targets                     | Missing |
| Whole-tree XML              | Bounded structural parts with fatal encoding, namespace, entity, declaration, depth, node, and duplicate expanded-attribute validation       | Missing |
| Streaming XML               | Incremental fatal UTF-8/UTF-16 decoding and arbitrary chunk splits; skipped branches still validated and counted                             | Missing |
| Markup Compatibility        | Versioned understood namespaces, deterministic choice/fallback, Ignorable and ProcessContent semantics, unknown non-ignorable rejection      | Missing |
| Selection syntax            | Sheet-name comparison, A1/range/whole-row/whole-column forms, invalid chart-sheet ranges, unknown/duplicate names                            | Missing |
| Selection work bounds       | Indexed membership, scanned versus returned accounting, dependencies outside emitted ranges, intersections returned unclipped                | Missing |

### Feature matrix

Every row below retains the class assigned by the reader plan.

| Area                         | Class      | Required completion evidence                                                                    | Status  |
| ---------------------------- | ---------- | ----------------------------------------------------------------------------------------------- | ------- |
| OPC package                  | Required   | Canonical discovery, ownership, relocated roots, malformed and adversarial packages             | Missing |
| Namespaces and compatibility | Required   | Strict/Transitional aliases and complete Markup Compatibility behavior                          | Missing |
| Workbook                     | Required   | Ordered sheets, properties, active sheet, visibility, date system, calculation state            | Missing |
| Calculation metadata         | Metadata   | Modes, versions, iteration/full-calc flags, chain metadata without execution                    | Missing |
| Sheet kinds                  | Required   | Worksheet/chart-sheet public union and fixtures                                                 | Missing |
| Sheet metadata               | Required   | Colors, dimensions, defaults, views, scenarios, selected/active state                           | Missing |
| Cells                        | Required   | Sparse authored cells, inferred refs, all value kinds, duplicates/order/bounds                  | Missing |
| Shared strings               | Required   | Plain/rich/phonetic text, whitespace, index and output-accounting boundaries                    | Missing |
| Formulas                     | Required   | Normal/shared/array/data-table/dynamic-array, caches, token-aware translation, no execution     | Missing |
| Modern cell metadata         | Required   | Rich values, cell images, checkboxes, spill and linked-data metadata under versioned namespaces | Missing |
| References                   | Required   | A1, quoted sheets, 3D, external, and structured references with seeded properties               | Missing |
| Defined names                | Required   | Workbook/sheet scopes, collisions, print areas/titles, hidden names                             | Missing |
| Styles                       | Required   | Fonts, fills, borders, alignment, formats, protection, named and differential styles            | Missing |
| Dates and times              | Required   | 1900/1904, serial 0/59/60/61, ISO dates, time, duration, date-time, locale/TZ independence      | Missing |
| Rows and columns             | Required   | Size, hidden, outline/collapse, defaults, overlapping column precedence                         | Missing |
| Views                        | Required   | Panes, selections, zoom, RTL, gridline/header state                                             | Missing |
| Merges                       | Required   | Ordered validated non-overlapping in-grid ranges and exact limits                               | Missing |
| Hyperlinks                   | Required   | Internal locations, allowlisted external protocols, no navigation/fetch                         | Missing |
| Notes and comments           | Required   | Legacy/threaded comments, authors/persons, visibility, missing relationships                    | Missing |
| Tables                       | Required   | Names, ranges, columns, headers/totals/formulas/styles, cardinality and ownership               | Missing |
| Filters and sorts            | Required   | Auto/custom/dynamic/top/color/icon filters and authored sort conditions                         | Missing |
| Data validation              | Required   | Types, operators, formulas, prompts/errors, multi-ranges and invalid sources                    | Missing |
| Conditional formatting       | Required   | All declared rule families, differential styles, priority/order/stop, extensions                | Missing |
| Drawings and images          | Required   | All anchors, geometry/crop/transform, relationship ownership, bounded media and URL lifecycle   | Missing |
| Charts                       | Required   | Common families, series formulas/caches, axes, titles, legends, colors/styles                   | Missing |
| Sparklines                   | Required   | Groups, data/location ranges, axes/colors, x14 namespaces                                       | Missing |
| Pivot tables                 | Required   | Definitions, fields, axes, filters, styles, bounded normalized model                            | Missing |
| Pivot caches                 | Metadata   | Definition always; records only by explicit mode with record/text limits                        | Missing |
| Slicers and timelines        | Metadata   | Cache links, ownership, and safe display metadata                                               | Missing |
| Print and layout             | Required   | Margins, orientation/paper/scale, repeating titles, breaks, headers/footers                     | Missing |
| Protection                   | Metadata   | Protected state and algorithm metadata without password/decryption claims                       | Missing |
| External links               | Metadata   | Safe redacted targets/formulas; no linked workbook access                                       | Missing |
| Connections/query tables     | Metadata   | Redacted safe metadata; no refresh, credentials, or connection strings                          | Missing |
| Active/embedded content      | Diagnostic | Recognize and omit/reject OLE, ActiveX, scripts, and executables by mode                        | Missing |
| Known extensions             | Required   | Namespace-specific normalized contracts and producer fixtures                                   | Missing |
| Unknown extensions           | Diagnostic | Stable safe omission or strict rejection without raw XML                                        | Missing |
| Document properties          | Required   | Core/app/custom typed values, malformed types, untrusted text                                   | Missing |

### Reader limits and failure contracts

The reader must publish and enforce every planned `XlsxResourceLimits` field:
package bytes and entries; expanded/per-part/XML/media bytes; XML depth/nodes;
worksheets, relationships, names; scanned and returned cells; row/column hard
bounds; shared strings, rich runs, text; formula sizes/groups; styles, ranges,
tables, links, validations, conditional rules, comments, drawings, charts, and
pivot records.

Completion evidence for each limit is: invalid configuration, below, exact,
one-over, structured limit metadata, fatal behavior in tolerant and strict
modes, and mutation coverage for comparisons and accounting. Limit dependency
validation and safe-integer arithmetic are separate required assertions.

Required parse codes are:

```text
invalid-package
invalid-document-structure
invalid-document-value
invalid-cell-reference
invalid-formula
invalid-selection
invalid-relationship-target
missing-required-part
resource-limit-exceeded
security-rejected-content
unsupported-feature
xml-parse-failed
xml-read-failed
```

No reader limit or diagnostic code currently has XLSX evidence.

## Round-trip contract audit

### Snapshot and fidelity

| Requirement                   | Required evidence                                                                                                      | Status  |
| ----------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ------- |
| Strict complete snapshot read | No tolerant omission, no selection, full source/security/capability inventory                                          | Missing |
| Portable JSON                 | Bounded standalone JSON with canonical Base64, source length/SHA-256, schema and profile versions                      | Missing |
| Canonical hashes              | Engine-independent base/state hashes including deterministic object keys; Node/browser equality                        | Missing |
| Stable snapshot keys          | Same exact source/profile yields identical keys; rename survival; deterministic added-object keys; collision rejection | Missing |
| Snapshot validation           | Plain-data schema/depth/object/string budgets, unknown-field rejection, immutability, source strict reparse            | Missing |
| Exact R0                      | JSON stringify/parse/validate/write returns byte-identical source and SHA for every accepted corpus source             | Missing |
| Edited R1                     | All copy-part uncompressed bytes identical and every changed/add/remove part declared                                  | Missing |
| Edited R2                     | Fresh strict reparse equals requested semantics and preserves unaffected supported semantics                           | Missing |
| Edited R3                     | Versioned Excel Windows/macOS, Calc, and applicable Sheets open/save evidence without repair                           | Missing |
| Capability manifest           | Machine-reviewable feature/operation/producer matrix with preservation-only/R1/R2/R3 states                            | Missing |
| Structured write errors       | Planned typed codes and bounded operation/object/part/feature/limit fields, never raw sensitive content                | Missing |

### Typed operation surface

Each operation needs schema, unique operation ID, optional canonical `ifMatch`,
ordered atomic replay, resource accounting, stable-key targeting, impact
closure, semantic preview/state hash, valid/invalid/missing/boundary/conflict/
inverse/concurrency tests, and an explicit capability-manifest level.

| Operation family                                                              | Status  |
| ----------------------------------------------------------------------------- | ------- |
| Set/clear cell content and set cell style                                     | Missing |
| Set/remove comment and hyperlink                                              | Missing |
| Set row and column properties                                                 | Missing |
| Insert/delete rows and columns                                                | Missing |
| Add/rename/delete/reorder worksheet                                           | Missing |
| Set/remove table, filter, validation, and conditional format                  | Missing |
| Set/remove drawing and image                                                  | Missing |
| Set chart and formula/cache policy                                            | Missing |
| Set pivot/slicer/timeline metadata                                            | Missing |
| Set print, protection, view, name, merge, and other advertised reader domains | Missing |

### Writer preservation and verification

| Requirement                     | Required evidence                                                                                                      | Status  |
| ------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ------- |
| Internal package graph          | Canonical parts, bytes/hashes, content types, owner-scoped relationships, opaque/active/signature/external markers     | Missing |
| Impact graph                    | Complete transitive semantic/package closure and conservative block outcomes                                           | Missing |
| Part dispositions               | Exactly one of copy/patch/rebuild/add/remove/block for every relevant part                                             | Missing |
| XML patch safety                | Recorded token spans, escaped non-overlapping patches, encoding preservation, full reparse and semantic check          | Missing |
| Stable allocation               | Preserve existing IDs; deterministic append; no canonical collisions, accidental renumbering, or unsafe orphan removal | Missing |
| Formula/reference transforms    | Token-aware supported rewrites, cached-result policy, chain/cache invalidation, recalculation flags, no evaluation     | Missing |
| Output graph validation         | Fresh reopen; complete ZIP/OPC/XML/relationship/content-type/security/limit checks                                     | Missing |
| Independent semantic validation | Fresh reader instance and literal operation-derived target comparison                                                  | Missing |
| Fidelity report                 | Bounded source/output hashes, dispositions, invalidated caches, recalculation and producer evidence                    | Missing |
| Deterministic ZIP edits         | Repeated/concurrent writes produce isolated deterministic state under the declared contract                            | Missing |

### Round-trip security and limits

Both `reject-active` and acknowledged `preserve-opaque` policies require public
tests. Active main formats remain rejected. Signed packages allow exact R0 but
edited output blocks without a separate signing/removal contract. External
data is never fetched, opaque conflicts fail closed, and no unsafe-preserve
escape hatch exists.

All planned `XlsxWriteLimits` need the same invalid/below/exact/one-over,
structured error, early-abort, and mutation evidence as reader limits. This
includes JSON, source, operation, dirty/patch, generated/output, formula
rewrite, reference/dependency, and validation-pass budgets.

No round-trip security mode or writer limit currently has XLSX evidence.

## Shared dependency audit

The following dependencies are already used by PPTX and may be consumed by
XLSX without changing their current contract:

| Dependency                     | XLSX use                                                | Shared-change risk                                                                                |
| ------------------------------ | ------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| `jszip`                        | Package open/write and independent fixture ZIPs         | XLSX must add actual expanded-byte accounting; JSZip metadata alone is insufficient.              |
| `saxes`                        | XML event parsing                                       | A new bounded streaming adapter is required; do not weaken current fatal validation.              |
| `common/archive/read-entry.ts` | Bounded expanded entry reads                            | Aggregate/package budgets and streaming abort may require a later generic extension.              |
| `common/binary/base64.ts`      | Portable source encode/decode and lexical length checks | Hashing remains separate; canonical Base64 behavior is already mutation-tested.                   |
| `common/xml/*`                 | Small structural parts and namespace-aware validation   | XLSX needs a streaming sibling with equivalent correctness and Markup Compatibility evidence.     |
| `common/opc/part-uri.ts`       | Owner-relative targets                                  | Package-wide canonical identities and percent/path ambiguity need stronger second-consumer tests. |
| `common/media/media-type.ts`   | Embedded image typing                                   | XLSX relationship ownership and media lifecycle remain format-owned.                              |
| `tinycolor2`                   | Potential DrawingML color normalization                 | No XLSX style or chart code may import PPTX domain modules.                                       |

No new runtime dependency is approved by this audit. A formula tokenizer,
formatting engine, or spreadsheet library would require the browser, bundle,
maintenance, security, license, state, and independent-oracle review specified
by the reader plan before installation.

## Deferred shared integration

The following files are deliberately deferred until the final integration
commit unless a proven shared primitive requires its own behavior-preserving
commit with full PPTX gates:

```text
src/index.ts
src/cli.ts
src/cli/*
package.json
tsup.config.ts
scripts/check-package.mjs
.github/workflows/*
README.md
docs/architecture.md
```

XLSX implementation and tests must import the format entry point directly
until that commit. `scripts/mutation-scope.mjs` is not deferred: every new XLSX
runtime target enters mutation scope with the commit that makes it behavioral.

## Required final evidence

Completion requires all of the following authoritative evidence:

1. Focused public and domain tests for every matrix row, failure state, and
   exact boundary.
2. Excel Windows/macOS, LibreOffice Calc, and Google Sheets corpus manifests
   with semantic assertions and exact R0 fingerprints.
3. Seeded ZIP/XML/reference/formula/operation fuzz and property suites with
   minimized regression cases.
4. Sequential/concurrent determinism, input immutability, selection scale,
   early abort, peak memory/RSS, elapsed-time, browser startup, and bundle-size
   baselines.
5. Node.js 20/22/24 and Chromium reader plus round-trip gates.
6. A complete mutation audit with no XLSX `Survived` or `NoCoverage` result and
   unchanged 100% thresholds/scope integrity.
7. Package ESM/CJS/declarations/subpath and CLI smoke tests.
8. A final requirement-by-requirement review of this audit in which no row is
   missing, indirect, or supported only by a narrower check.

Until all rows are proven, README and architecture capability text must not
describe XLSX reader or round-trip support as complete.
