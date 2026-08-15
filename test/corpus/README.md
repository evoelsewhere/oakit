# PPTX reliability corpus

The corpus is downloaded into `.cache/pptx-corpus`; binary documents are not
redistributed in the repository. Every download has provenance and a pinned
integrity fingerprint in `pptx-manifest.json`.

- PowerPoint-produced documents come from LibreOffice's MPL-2.0 regression
  repository. Producer metadata was verified from each package where present.
- The LibreOffice case is generated locally by a real headless LibreOffice
  round-trip of the pinned PowerPoint fixture.
- The large Google Slides case is fetched directly from a publicly indexed
  Google Slides PPTX export URL and is only included by the large corpus gate.
  Google rebuilds that ZIP and its media on every export, so this entry pins the
  stable SHA-256 of its ordered slide text plus a maximum download size instead
  of an unstable whole-file digest.

Run `pnpm test:corpus` for the curated gate or `pnpm test:corpus:large` for the
large Google Slides export as well. A fingerprint mismatch fails closed so a
changed remote document cannot silently alter the test corpus.
