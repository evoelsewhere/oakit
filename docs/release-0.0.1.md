# Release 0.0.1 readiness

OAKit `0.0.1` is not released merely because an individual gate is green. The
release is ready only when every required item below is proven for one immutable
production source tree and the final metadata revision passes package and CI
verification. Evidence files cannot contain the hash of the commit that contains
them, so the release binds producer and Reliability runs through the shared
`src/` Git tree `4b7c749f58f908967b9868716bddd78e05a0fbcd`.

## Capability claims

- `pptx-create-text-v1`: effective `C3` for the declared PowerPoint,
  LibreOffice Impress, and Google Slides producer matrix.
- `pptx-roundtrip-text-v1`: effective `R3` for supported plain-text replacement
  and text move/resize/rotate/flip operations across the same matrix.
- Portable unchanged hand-off: byte-exact `R0`.
- Preview: self-contained SVG and Node PNG without an Office runtime, with
  explicit approximation warnings.

These are versioned profile claims, not blanket support for arbitrary PPTX
creation, arbitrary element editing, or pixel-identical rendering.

## Required release gates

- [x] Producer matrix manifests retained for PowerPoint, LibreOffice, and Google
      Slides.
- [x] Real-world transient corpus strict-parses and renders without Office before
      and after controlled Google Slides import/export.
- [x] Focused mutation modules pass at 100% without excluded mutants.
- [x] Release mutation architecture includes forced dynamic and static mutants.
- [x] Node 20/22/24, browser, fuzz, package, CLI, and resource-limit gates exist.
- [x] Refresh the controlled Google credential and regenerate producer/corpus
      evidence for the release source tree in
      [run 32045412714](https://github.com/evoelsewhere/oakit/actions/runs/32045412714).
- [x] Run the complete Reliability workflow for the same source tree and audit
      the merged mutation artifact with zero missed mutants in
      [run 32049829830](https://github.com/evoelsewhere/oakit/actions/runs/32049829830):
      21,184 mutants, 16,734 killed, 4,450 compile errors, and zero survived,
      no-coverage, timeout, or runtime-error results.
- [x] Require a fully green CI run on the Reliability revision; all 11 jobs
      passed in
      [run 32049822932](https://github.com/evoelsewhere/oakit/actions/runs/32049822932).
- [x] Replace the `Unreleased` changelog heading with `0.0.1` and its release
      date.
- [x] Change `package.json` from `0.0.0` to `0.0.1` after attaching the preceding
      evidence. The pnpm v9 lockfile does not store the root package version.
- [x] Pack and install the release tarball, verify ESM/CJS/subpath exports and
      `oakit --version` output, and require a fully green CI run on the metadata
      revision. Local packed-consumer smoke passed and all 11 CI jobs passed in
      [run 32053915220](https://github.com/evoelsewhere/oakit/actions/runs/32053915220).
- [x] Verify the release candidate changes no production source from the audited
      tree; `83f1cebf0dd1429a1f8305389681cc50e010b038` retains source tree
      `4b7c749f58f908967b9868716bddd78e05a0fbcd`.
- [x] Create tag `v0.0.1` from the final docs-only attestation descendant at
      `42a7483fd101c2374cf6afe4e656f2da19787679`.
- [x] Publish the [GitHub release](https://github.com/evoelsewhere/oakit/releases/tag/v0.0.1)
      and complete provenance-enabled npm
      [run 32054316824](https://github.com/evoelsewhere/oakit/actions/runs/32054316824).

## Published artifact note

npm 11 removed the `oakit` bin mapping while normalizing the published `0.0.1`
manifest because its path started with `./`. The ESM, CJS, and subpath APIs remain
available, but the installed command is not claimed for this immutable package
version. The release notes disclose the issue, the manifest now uses
`dist/cli.js`, and package smoke now executes `node_modules/.bin/oakit` directly.
That fix is released as `0.0.2`; the `v0.0.1` tag is not moved or rewritten.

## Evidence

- Text-profile producer matrix: [`evidence/0.0.1/producer-matrix.json`](evidence/0.0.1/producer-matrix.json)
- Producer notes and mutation audit: [`evidence/0.0.1/README.md`](evidence/0.0.1/README.md)
- Real-world corpus evidence: [`evidence/0.0.1/slidesmania/evidence.json`](evidence/0.0.1/slidesmania/evidence.json)
- Current corpus audit image: [`evidence/0.0.1/slidesmania/producer-audit.png`](evidence/0.0.1/slidesmania/producer-audit.png)
- Machine-readable release gates: [`evidence/0.0.1/release-gates.json`](evidence/0.0.1/release-gates.json)
