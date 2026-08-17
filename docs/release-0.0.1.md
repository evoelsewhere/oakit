# Release 0.0.1 readiness

OAKit `0.0.1` is not released merely because an individual gate is green. The
release is ready only when every required item below is proven on the exact
release revision.

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
- [ ] Refresh the controlled Google credential and regenerate producer/corpus
      evidence on the final release revision.
- [ ] Run the complete Reliability workflow on the final release revision and
      audit the merged mutation artifact with zero missed mutants.
- [ ] Require a fully green CI run on the final release revision; infrastructure
      setup failures must be rerun rather than treated as product failures or
      ignored.
- [ ] Replace the `Unreleased` changelog heading with `0.0.1` and its release
      date.
- [ ] Change `package.json` and lockfile metadata from `0.0.0` to `0.0.1` only
      after the preceding evidence is attached to the same revision.
- [ ] Pack and install the release tarball, verify ESM/CJS/subpath exports and
      CLI version output, then create tag `v0.0.1`.
- [ ] Publish the GitHub release to trigger provenance-enabled npm publishing.

## Evidence

- Text-profile producer matrix: [`evidence/0.0.1/producer-matrix.json`](evidence/0.0.1/producer-matrix.json)
- Producer notes and mutation audit: [`evidence/0.0.1/README.md`](evidence/0.0.1/README.md)
- Real-world corpus evidence: [`evidence/0.0.1/slidesmania/evidence.json`](evidence/0.0.1/slidesmania/evidence.json)
- Current corpus audit image: [`evidence/0.0.1/slidesmania/producer-audit.png`](evidence/0.0.1/slidesmania/producer-audit.png)
