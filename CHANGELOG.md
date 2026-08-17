# Changelog

All notable changes to OAKit are documented here.

## Unreleased

### Added

- bounded PowerPoint parsing with strict diagnostics and resource limits;
- portable JSON snapshots with byte-exact `R0` restore and integrity binding;
- producer-verified `C3` source-free creation for `pptx-create-text-v1`;
- producer-verified `R3` plain-text and text-transform editing for
  `pptx-roundtrip-text-v1`;
- Office-free SVG and PNG rendering in Node.js and browser-compatible SVG;
- native preview rendering for rich text, authored fonts, connectors, linear
  and radial gradients, image crops, and table dimensions, merges, borders,
  and vertical text alignment;
- real-world transient SlidesMania and controlled producer reliability gates;
- focused mutation modules for pull requests and forced full/static mutation
  gates for release validation.

### Security

- fail-closed handling for unsafe relationships, XML, ZIP/resource limits,
  signatures, protected packages, unsupported edit ownership, and malformed
  portable state.
