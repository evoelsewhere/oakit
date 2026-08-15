# Development instructions

These instructions apply to the entire repository. Follow them together with
the public contracts in `README.md` and the boundaries in
`docs/architecture.md`.

## Mission and scope

`office2json` converts Office Open XML packages into normalized application
models. The current production capability is PPTX-to-JSON. Excel, Word, and
JSON-to-Office are separate future format milestones; do not imply that they
exist until their public APIs, implementations, and tests are present.

Optimize changes for:

- correct and deterministic public output;
- safe handling of untrusted ZIP and XML input;
- PowerPoint, LibreOffice, Google Slides, Node.js, and browser compatibility;
- isolated format domains that can evolve without leaking raw OOXML;
- small, reviewable changes backed by independent evidence.

## Read before changing code

For any non-trivial change:

1. Read the relevant sections of `README.md` and `docs/architecture.md`.
2. Locate the public type and the module that owns the behavior.
3. Search existing tests and fixtures before creating a new abstraction.
4. Check `git status` and preserve unrelated user changes.
5. State the public invariant that the change must preserve or introduce.

Treat implementation comments and historical behavior as secondary evidence.
The public API, OOXML ownership rules, tests, and producer documents are the
primary evidence.

## Architecture boundaries

Preserve this dependency direction:

```text
root API
  -> format public API and types
    -> format parser/orchestrator
      -> format internal domain modules
        -> shared format-neutral primitives
```

- `src/index.ts` only exposes supported format entry points.
- `src/formats/pptx/index.ts` and `types.ts` define the supported PPTX API.
- `src/formats/pptx/parser.ts` owns package and slide orchestration.
- `src/formats/pptx/internal/*` owns PowerPoint-specific domain behavior.
- `src/common/*` must remain independent of PowerPoint node paths and
  inheritance rules.
- Raw XML nodes, relationship IDs, parser contexts, and caches must not escape
  into the public document model.
- New Office formats must be siblings under `src/formats`, not branches inside
  the PPTX parser.
- Reader and writer orchestration must remain separate. Do not implement a
  writer by reversing tolerant reader steps.

Keep mutable parsing state scoped to one parse operation. Caches may be shared
between slides of the same document, but never between documents or concurrent
calls.

## OOXML ownership rules

- Resolve a relationship relative to the part that owns its `.rels` file.
- Resolve internal targets through the OPC URI helpers; never concatenate
  package paths ad hoc.
- Preserve external targets only when `TargetMode="External"` is explicit.
- Apply inheritance from the most specific source to the least specific source
  and make fallback behavior explicit.
- Normalize units, colors, text, paths, and chart data before returning them.
- Preserve authored document order. Do not infer order from object keys or ZIP
  entry enumeration.
- Treat missing optional parts differently from malformed required parts.

## Correctness and security invariants

These rules are mandatory for every input path:

- Reject ZIP traversal, unsafe relationship targets, malformed XML, invalid
  UTF encodings, duplicate expanded attributes, and invalid numeric values.
- Enforce compressed input, entry count, expanded bytes, per-part bytes, XML
  depth, XML node count, media size, and slide count limits.
- Resource-limit violations are fatal in both strict and tolerant modes.
- Tolerant mode may recover from malformed optional content only by recording a
  diagnostic and omitting the unsafe result.
- Strict mode must throw the typed public parse error carrying its diagnostic.
- Never return `NaN`, `Infinity`, negative dimensions caused by malformed
  input, or partially trusted relationship paths.
- Escape HTML text and attributes, serialize CSS values safely, and allow only
  explicitly supported hyperlink protocols.
- Do not execute macros, scripts, media, hyperlinks, or external resources.
- Do not fetch external relationships during parsing.
- Browser object URLs belong to the caller; tests that create them must revoke
  them.

Do not weaken a validation rule merely to accept a corpus document. First
determine whether the document is valid, whether the rule is overly strict, or
whether the producer requires a documented compatibility branch.

## Hard implementation rules

The rules in this section are blocking. An implementation is not complete when
one of them is violated, even if the default test suite passes.

### Evidence before production code

1. Every bug fix must begin with a regression test that fails for the expected
   reason before production code is changed.
2. Every behavior change must identify its public invariant, failing input,
   expected output or diagnostic, and strict/tolerant behavior.
3. A test must assert semantic output, a typed diagnostic, or a security
   boundary. “Does not throw” is not sufficient evidence.
4. A minimized fuzz or mutation counterexample must remain as a deterministic
   regression test.
5. Do not update an expected value merely to make a failing test pass. Explain
   why the previous contract was incorrect and update the public contract when
   necessary.

### Scope and ownership

6. Change only the module that owns the behavior. Do not include opportunistic
   cleanup or unrelated formatting.
7. Do not combine a feature, broad refactor, dependency update, and formatting
   pass in one logical change.
8. Do not change a public type to make an internal implementation easier.
9. A breaking public API correction requires a migration note and explicit
   tests for the affected contract.
10. Code may move into `src/common` only when it is format-neutral and has two
    independent consumers, or when the format-neutral contract is otherwise
    demonstrated by focused tests.

### Type and state safety

11. Do not add `any`, `@ts-ignore`, `@ts-nocheck`, broad lint suppression, or an
    unchecked double cast to bypass a design problem.
12. A type assertion at an external boundary requires runtime validation before
    the value is trusted.
13. Handle discriminated unions exhaustively and use `never` for unreachable
    states.
14. Do not use a non-null assertion for caller, ZIP, relationship, or XML data.
15. Missing, malformed, unsupported, and intentionally empty values must remain
    distinguishable states.
16. Do not mutate caller-owned bytes, options, or objects.
17. The same input and options must produce the same result during sequential
    and concurrent parsing.

### Numeric and fallback behavior

18. Validate every numeric attribute in this order:

    ```text
    present -> parse -> finite -> safe/range check -> unit conversion
    ```

19. Do not return a value produced by `Number`, `parseInt`, or `parseFloat`
    before validating finiteness, sign, integer requirements, and range.
20. Every numeric or size limit needs tests for exactly-at-limit,
    one-over-limit, and invalid input.
21. A fallback is allowed for absent or explicitly unsupported data. Existing
    but malformed data must produce a diagnostic or rejection instead of a
    silent fallback.
22. Public output must not contain non-finite numbers, unnormalized package
    paths, raw relationship IDs, or partially parsed unsafe values.

### ZIP, XML, and relationship boundaries

23. Never construct an internal package target with manual string
    concatenation. Use the OPC resolver owned by the source part.
24. Never preserve an external URI without explicit external target mode.
25. Never fetch, execute, navigate to, or otherwise dereference an external
    relationship while parsing.
26. Enforce resource limits before expansion when metadata is trustworthy and
    during expansion regardless of metadata.
27. Validate XML encoding, lexical structure, entities, nesting, and node limits
    before trusting the dependency parser result.
28. Do not remove a validation rule to accommodate one producer fixture. Prove
    that the input is valid and add a documented compatibility path with its
    own regression test.

### Error handling

29. Do not add an empty `catch` block or catch an error only to return `null`,
    an empty object, or a default value.
30. A caught error must be converted to a typed diagnostic, enriched and
    rethrown, or recovered through an explicitly tested tolerant-mode rule.
31. Do not branch on exception message text.
32. Preserve diagnostic code, part, severity, cause, and resource-limit metadata
    when that information exists.
33. Security tests must assert error type and diagnostic code. Message text may
    be asserted additionally, but never as the only contract.

### Test independence

34. Public behavior tests must call the public API.
35. A fixture builder must not call the production helper whose result the test
    is intended to verify.
36. Expected values must not be computed with the same algorithm as production
    code.
37. A snapshot cannot be the only assertion for security, numeric limits,
    relationships, or diagnostics.
38. Every new parser branch needs valid, malformed, and missing-optional cases.
    Add strict/tolerant cases whenever recovery behavior differs.
39. Tests that create browser object URLs must read back the expected bytes and
    revoke every URL in cleanup.
40. Property tests must use a recorded seed and bounded runs. A failure report
    must retain the seed and minimized counterexample.

### Mutation quality

41. Changed lines must introduce no new `Survived` or `NoCoverage` mutants.
42. The mutation score of every changed target file must not decrease.
43. Do not improve mutation score by shrinking target files, excluding a
    meaningful mutator, reducing test scope, or asserting private implementation
    details.
44. Refactor equivalent mutants into simpler observable logic. Ignore a mutant
    only with a line-specific explanation proving equivalence.
45. Investigate timeout mutants as possible infinite loops or unbounded work;
    do not accept them automatically as adequate detection.

### Complexity and dependencies

46. A function must not combine ZIP I/O, XML validation, normalization, and
    public model construction. Split these phases into independently testable
    units.
47. Refactor a function when it exceeds three meaningful nesting levels or has
    several independent failure modes that cannot be tested in isolation.
48. A new runtime dependency requires a written rationale covering browser
    compatibility, bundle size, maintenance, security history, and why existing
    code is insufficient.
49. Do not update dependencies outside the task scope.
50. Do not add Node-only APIs to any source path reachable from the browser
    entry point.

### Required proof at hand-off

51. Run the focused regression test and every gate required by the changed
    boundary.
52. Report the exact commands and pass counts; do not say “tests pass” without
    execution evidence.
53. Report remaining uncovered or surviving mutation behavior explicitly.
54. Verify that the working tree contains no unrelated edits before staging or
    handing off the change.
55. Do not commit implementation while a focused regression test or required
    gate is failing.

## Implementation workflow

Use this sequence for fixes and fidelity work:

1. Reduce the behavior to the smallest independently constructed OOXML package
   that reproduces it.
2. Add a test against the public API and confirm that it fails for the expected
   reason.
3. Change the narrowest owning module.
4. Add exact boundary pairs where relevant: equal-to-limit and one-over-limit,
   valid and invalid, strict and tolerant, internal and external.
5. Run the focused tests during iteration.
6. Run the complete required and relevant extended gates.
7. Update public types, README, and architecture documentation when their
   contracts or boundaries changed.

Avoid speculative generalization. Move code into `src/common` only after its
contract is demonstrably format-neutral.

## Test policy

Tests must make failures independently observable. Do not reproduce production
logic in the expected value and do not assert only that parsing did not throw.

Prefer these layers:

1. A black-box test through exports from `src` for public behavior.
2. A focused primitive test for a genuinely shared internal contract.
3. A real-producer corpus assertion when producer compatibility is the
   behavior under test.
4. A seeded property test when the input space is larger than useful examples.

Use `test/black-box/pptx-package.ts` to build reviewable packages in memory.
Assert the normalized document, typed diagnostic, or security boundary—not raw
intermediate XML objects.

For bug fixes, retain the minimized counterexample as a deterministic
regression test. Property tests must use recorded seeds and bounded run counts
so failures are reproducible in CI.

Corpus binaries are downloaded into `.cache`; do not commit them. Add source,
producer, license/provenance, integrity fingerprint, size bound, and semantic
expectations to `test/corpus/pptx-manifest.json`. A corpus test must assert
meaningful content counts so an empty document cannot pass.

## Quality gates

Install and use the pinned package manager:

```bash
pnpm install
```

During development, run the smallest relevant test first:

```bash
pnpm exec vitest run path/to/test.ts
```

Every completed code change must pass:

```bash
pnpm check
```

`pnpm check` covers formatting, linting, strict type checking, the deterministic
Vitest suite, declarations, ESM, and CommonJS builds.

Run additional gates when their boundary is touched:

```bash
pnpm test:browser       # browser input, Blob, media, object URL, bundling
pnpm test:corpus        # PowerPoint and LibreOffice compatibility
pnpm test:corpus:large  # includes the large Google Slides export
pnpm test:mutation      # assertion quality for core safety logic
```

Changes to ZIP, XML, OPC relationships, sanitization, resource limits, public
input types, or media loading require the mutation gate. Changes affecting
runtime-neutral code require the browser gate. Parser fidelity changes require
the curated corpus; cross-producer or scaling changes require the large corpus.

Do not make mutation results look better by shrinking the target files,
excluding meaningful mutators, or asserting private implementation details.
Convert surviving mutants into boundary tests or refactor equivalent branches
into simpler observable code. The configured mutation break threshold is a
regression floor, not the target quality level.

## Code conventions

- Use ESM and type-only imports where applicable.
- Preserve strict TypeScript settings; do not introduce `any`, unchecked casts,
  broad suppression comments, or disabled lint rules to bypass a design issue.
- Prefer explicit guards and discriminated unions over assertions.
- Keep functions small enough that validation, traversal, normalization, and
  serialization can be tested separately.
- Use safe integer parsing for OOXML numeric attributes and validate before
  unit conversion.
- Keep production code runtime-neutral. Do not add filesystem or other Node-only
  APIs to parser paths used by browsers.
- Do not edit generated `dist` files.
- Reuse existing dependencies and helpers before adding a dependency. New
  runtime dependencies require a clear size, security, and browser rationale.

Preserve established public spellings for compatibility, even when they are
historical mistakes. A correction requires a deliberate migration plan and
tests.

## Documentation and public API changes

When a public element, option, diagnostic, or return value changes, update in
the same logical change:

- the public discriminated unions and exported types;
- the parser and owning domain module;
- black-box fixtures and consumer-facing assertions;
- README examples or tables;
- `docs/architecture.md` if ownership, lifecycle, or pipeline behavior changed;
- generated declaration verification through `pnpm build`.

Do not claim planned formats or reverse conversion as completed behavior.

## Git and commits

- Keep commits small, atomic, and buildable.
- Use ordinary intent-based subjects such as `fix:`, `test:`, `docs:`, `ci:`,
  `refactor:`, or `chore:`.
- Do not put planning priority labels such as `P0` or `P1` in commit subjects.
- Separate fixture/test additions, implementation fixes, infrastructure, and
  documentation when each is independently meaningful.
- Never stage, rewrite, discard, or reformat unrelated user changes.
- Do not rewrite history, force-push, or push unless the user explicitly asks.

## Definition of done

A change is complete only when:

- its public invariant is stated and tested;
- malformed and boundary input behavior is explicit;
- relevant Node, browser, corpus, fuzz, and mutation gates pass;
- no unrelated working-tree changes were included;
- documentation matches actual capabilities;
- commits describe behavior rather than internal activity or planning priority.
