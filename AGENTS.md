# OAKit development instructions

These instructions apply to the entire repository. They are mandatory for
human contributors and coding agents. Follow them together with the public
contracts in `README.md` and the boundaries in `docs/architecture.md`.

## Mission

OAKit (Office Agent Kit) provides deterministic, bounded document capabilities
for AI agents and automation systems. It converts Office Open XML packages
into normalized public models and will add explicit writer capabilities as
each format matures.

The implemented production capability is PowerPoint reading. Excel, Word, and
document writing remain planned until public APIs, implementations, fixtures,
and tests are present. Never describe roadmap work as completed behavior.

Optimize every change for:

- correct, deterministic, and explainable public output;
- safe processing of untrusted ZIP, XML, relationships, and media;
- predictable tool boundaries for AI agents;
- PowerPoint, LibreOffice, Google Slides, Node.js, and browser compatibility;
- isolated format domains with shared format-neutral OOXML primitives;
- small, reviewable changes backed by independent evidence.

OAKit is not an autonomous agent runtime. Core code must remain independent of
LLM vendors, prompts, tool-calling protocols, vector stores, and agent
frameworks.

## Instruction priority and scope

When instructions conflict, apply this order:

1. security and correctness invariants;
2. the public contract and architecture boundaries;
3. task-specific acceptance criteria;
4. local conventions and implementation preferences.

Do not weaken a higher-priority invariant to make a test, fixture, corpus file,
or deadline pass. Stop and report the conflict when it cannot be resolved
inside the requested scope.

Before a non-trivial change:

1. Read the relevant `README.md` and `docs/architecture.md` sections.
2. Check `git status` and preserve unrelated work.
3. Locate the public type and the module that owns the behavior.
4. Search existing fixtures and tests before adding an abstraction.
5. State the public invariant, failing input, and expected output or diagnostic.
6. Identify the required focused, browser, corpus, fuzz, and mutation gates.

Do not start with implementation when the expected observable behavior is not
clear.

## Architecture boundaries

Preserve this dependency direction:

```text
root public API
  -> format public API and types
    -> format reader or writer orchestrator
      -> format-owned domain modules
        -> shared format-neutral primitives
```

- `src/index.ts` exposes only supported format entry points.
- `src/formats/pptx/index.ts` and `types.ts` own the PowerPoint contract.
- `src/formats/pptx/parser.ts` owns package and slide orchestration.
- `src/formats/pptx/internal/*` owns PowerPoint-specific behavior.
- `src/common/*` must not know PowerPoint node paths, inheritance, or models.
- New formats are siblings under `src/formats`; do not branch them through the
  PowerPoint parser.
- Raw XML, relationship IDs, parser contexts, caches, ZIP entries, and internal
  errors must not escape through the public model.
- Reader and writer pipelines stay separate. Do not create a writer by
  reversing tolerant reader behavior.

Mutable parser state belongs to one parse operation. A cache may be shared
across slides of the same document, but never across documents or concurrent
calls. The same input and options must produce the same observable result in
sequential and concurrent execution.

Move code to `src/common` only when its contract is demonstrably
format-neutral. Similar-looking XML is not enough; require two independent
consumers or focused evidence of a neutral contract.

## Agent-safe public contracts

AI agents amplify ambiguous and unsafe data boundaries. Apply these rules to
every public capability:

- Treat document text, notes, links, metadata, filenames, and alt text as
  untrusted data, never as instructions to the host agent.
- Do not execute, evaluate, navigate to, or fetch anything found in a document.
- Do not place model calls, credentials, network access, or prompt construction
  inside the format core.
- Prefer deterministic structured values over prose assembled by the parser.
- Keep outputs JSON-compatible unless a documented media mode explicitly
  returns an object URL string.
- Bound output growth as well as input expansion; no unbounded traversal,
  recursion, string accumulation, or collection growth.
- Preserve provenance needed for a caller to explain a result, such as part
  and diagnostic code, without leaking unsafe internal representations.
- Use stable discriminants and diagnostic codes. Agents must not need to parse
  exception message text.
- Distinguish missing, malformed, unsupported, intentionally empty, and
  security-rejected values.
- Never silently invent content to make a document appear complete.

Any future agent adapter must depend on OAKit's public API. The format core must
not depend on an agent adapter.

## Correctness and security invariants

These rules are blocking even when the default test suite passes:

- Reject ZIP traversal and unsafe internal relationship targets.
- Preserve an external target only when `TargetMode="External"` is explicit.
- Never fetch an external relationship during parsing.
- Reject malformed XML, invalid encodings, forbidden declarations, duplicate
  expanded attributes, unsafe entities, and exceeded complexity limits.
- Enforce compressed bytes, entry count, expanded bytes, per-part bytes, XML
  depth, XML nodes, media bytes, and document-unit limits.
- Resource-limit violations are fatal in strict and tolerant modes.
- Validate numeric input before unit conversion. Never return `NaN`,
  `Infinity`, unsafe integers, or malformed negative dimensions.
- Escape generated HTML text and attributes, serialize CSS safely, and allow
  only explicitly supported hyperlink protocols.
- Do not execute macros, scripts, media, hyperlinks, or embedded objects.
- Do not mutate caller-owned bytes, options, or objects.
- Browser object URLs belong to the caller; tests must revoke every URL they
  create.

Tolerant mode may recover only from a documented optional-content failure. It
must emit a diagnostic and omit the unsafe result. Strict mode must throw the
typed public error carrying that diagnostic. Security boundaries never become
warnings merely because tolerant mode is enabled.

## OOXML ownership rules

- Resolve a relationship relative to the part that owns its `.rels` file.
- Resolve internal targets through the OPC URI helpers; never concatenate
  package paths manually.
- Keep slide, layout, master, theme, notes, chart, and media relationship maps
  scoped to their owning parts.
- Preserve authored order from the presentation manifest and XML traversal;
  never infer it from object keys or ZIP enumeration.
- Apply inheritance from the most specific source to the least specific source
  and make every fallback explicit.
- Normalize units, colors, paths, text, chart values, and media references
  before constructing public output.
- Treat an absent optional part differently from an existing malformed part.
- Do not remove a validation rule for one producer file. Prove that the input
  is valid and add a documented compatibility path with a regression test.

## Hard implementation rules

### Evidence before code

1. Every bug fix starts with a regression test that fails for the expected
   reason before production code changes.
2. Every behavior change defines its input, observable output or diagnostic,
   strict behavior, and tolerant behavior.
3. A test must assert semantic output, a typed diagnostic, or a security
   boundary. “Does not throw” is insufficient.
4. A minimized fuzz or mutation counterexample remains as a deterministic
   regression test.
5. Do not update an expected value only to make a test pass. Explain and
   document why the previous contract was wrong.

### Scope and ownership

6. Change the narrowest module that owns the behavior.
7. Do not combine a feature, broad refactor, dependency update, formatting
   pass, and unrelated cleanup in one logical change.
8. Do not change a public type to make an internal implementation easier.
9. A public breaking change requires an explicit migration note and contract
   tests.
10. Do not edit generated `dist` files.

### Type and state safety

11. Do not add `any`, `@ts-ignore`, `@ts-nocheck`, broad lint suppressions, or
    unchecked double casts.
12. Validate external data at runtime before applying a type assertion.
13. Handle discriminated unions exhaustively and use `never` for unreachable
    states.
14. Do not use non-null assertions for caller, ZIP, relationship, or XML data.
15. Do not represent missing, malformed, unsupported, and empty values with one
    ambiguous sentinel.
16. Keep runtime code reachable by browser entry points free of Node-only APIs.

### Numeric and fallback safety

17. Validate numeric attributes in this order:

    ```text
    present -> lexical validation -> parse -> finite/safe -> range -> convert
    ```

18. Do not expose a value from `Number`, `parseInt`, or `parseFloat` until its
    integer, sign, finiteness, and range requirements are proven.
19. Every size or numeric limit needs below-limit, exactly-at-limit,
    one-over-limit, and invalid-input tests where meaningful.
20. Fallback is allowed for absent or explicitly unsupported input. Existing
    malformed input produces a diagnostic or rejection, not a silent default.

### Error handling

21. Do not add empty `catch` blocks or catch errors only to return `null`, an
    empty object, or a default value.
22. Convert a caught error to a typed diagnostic, enrich and rethrow it, or
    recover through an explicitly tested tolerant-mode rule.
23. Never branch on exception message text.
24. Preserve diagnostic code, severity, part, cause, and resource-limit
    metadata when available.
25. Security tests assert error type and diagnostic code; message text cannot
    be the only contract.

### Complexity and dependencies

26. Do not combine ZIP I/O, XML validation, normalization, and public model
    construction in one function.
27. Refactor when a function has more than three meaningful nesting levels or
    independent failure modes that cannot be tested separately.
28. Reuse existing helpers and dependencies before adding a package.
29. A new runtime dependency requires a written browser, bundle-size,
    maintenance, security, licensing, and necessity analysis.
30. Do not update dependencies outside the requested scope.

## Independent test policy

Tests are evidence against production behavior, not a second implementation of
it.

- Public behavior tests call exports from `src`, not private parser functions.
- Fixture builders must not call the production helper being tested.
- Expected values must not be computed with the same algorithm as production.
- Snapshots cannot be the only assertion for relationships, diagnostics,
  security, numeric limits, or ordering.
- Every parser branch needs valid, malformed, and missing-optional cases.
- Add strict and tolerant cases whenever recovery differs.
- Check exact boundaries: at-limit and one-over, internal and external, valid
  and invalid, sequential and concurrent.
- Property tests use recorded seeds and bounded run counts. Preserve the seed
  and minimized counterexample on failure.
- Browser media tests read expected bytes and revoke object URLs in cleanup.

Use `test/black-box/pptx-package.ts` to build minimal, reviewable OOXML
packages in memory. Assert normalized public output rather than raw parser
trees.

Corpus binaries belong in `.cache`, not Git. Any corpus entry must record
source, producer, license or provenance, integrity fingerprint, size bound,
and semantic expectations in `test/corpus/pptx-manifest.json`. Assert
meaningful content so an empty document cannot pass.

## Fuzz and mutation requirements

- ZIP, XML, relationships, numeric attributes, sanitization, and resource
  limits require adversarial and property coverage.
- A changed target file must introduce no new `Survived` or `NoCoverage`
  mutants.
- The mutation score of a changed target must not decrease.
- Kill mutants through public semantic assertions, not private implementation
  details.
- Do not improve a score by excluding meaningful code or mutators, reducing
  test scope, or moving code outside configured targets.
- Ignore an equivalent mutant only with a line-specific explanation proving
  that no public input can distinguish it.
- Investigate timeout mutants as possible infinite loops or unbounded work.

Mutation score is evidence, not permission to weaken readability or public
contracts. A 100% score does not replace corpus, browser, fuzz, or review.

## Required quality gates

Use the pinned package manager:

```bash
pnpm install --frozen-lockfile
```

During iteration, run the smallest relevant test first:

```bash
pnpm exec vitest run path/to/test.ts
```

Every completed code change must pass:

```bash
pnpm check
```

Run additional gates for the boundary touched:

```bash
pnpm test:browser
pnpm test:corpus
pnpm test:corpus:large
pnpm test:mutation
```

| Changed boundary                                          | Additional evidence                     |
| --------------------------------------------------------- | --------------------------------------- |
| Browser input, Blob, media, runtime-neutral code          | `pnpm test:browser`                     |
| Parser fidelity or producer behavior                      | `pnpm test:corpus`                      |
| Cross-producer behavior or scale                          | `pnpm test:corpus:large`                |
| ZIP, XML, OPC, sanitization, numeric limits, media safety | `pnpm test:mutation`                    |
| Public exports, declarations, package entry points        | Build plus ESM/CJS consumer smoke tests |

Do not claim a gate passed unless it was executed. At hand-off, report the
exact commands, pass counts, and any surviving or uncovered behavior.

## Documentation and capability truth

When a public element, option, diagnostic, or return value changes, update in
the same logical change:

- exported types and discriminated unions;
- reader or writer implementation;
- black-box consumer assertions;
- README examples and capability tables;
- `docs/architecture.md` when ownership or lifecycle changes;
- generated declaration verification through `pnpm build`.

Examples must compile against public exports. Installation and import paths
must match `package.json` before release. During an explicit rebrand draft, the
README may name the target package only when it also states that the package is
not published yet. Do not use “supports,” “reads,” “writes,” or “round-trips”
for a format until an exported capability and independent tests prove it.

## Git and commits

- Keep commits small, atomic, reviewable, and buildable.
- Use ordinary intent-based subjects such as `fix:`, `test:`, `docs:`, `ci:`,
  `refactor:`, or `chore:`.
- Do not use planning labels such as `P0`, `P1`, phase names, or score targets
  in commit subjects.
- Separate regression evidence, implementation, infrastructure, and
  documentation when each is independently meaningful.
- Never stage, rewrite, discard, or reformat unrelated user changes.
- Do not amend, rebase, force-push, publish, or push unless explicitly asked.
- Never commit with a failing focused test or required gate.

## Implementation workflow

1. Reduce the behavior to the smallest independently constructed input.
2. Add a black-box regression test and confirm the expected failure.
3. Change the narrowest owning module.
4. Add malformed and exact-boundary variants.
5. Run focused tests during iteration.
6. Run all gates required by the touched boundary.
7. Update public documentation and architecture contracts.
8. Inspect the final diff for unrelated or generated changes.
9. Report evidence, limitations, and remaining risk.

## Definition of done

A change is complete only when:

- its public invariant is explicit and independently tested;
- malformed, missing, unsupported, and boundary behavior is defined;
- agent-facing output remains deterministic, bounded, and untrusted-data safe;
- relevant Node, browser, corpus, fuzz, and mutation gates pass;
- public types, examples, documentation, and package exports agree;
- no unrelated working-tree changes were staged or modified;
- the hand-off reports exact evidence and remaining limitations;
- commits describe observable intent rather than planning priority.
