# Changelog

All notable changes to OAKit are documented here.

## 0.0.2 - 2026-08-18

### Fixed

- preserve the installed `oakit` command when npm 11 normalizes the published
  package manifest;
- execute `node_modules/.bin/oakit` in packed-consumer smoke tests so missing bin
  mappings fail before publication.

## 0.0.1 - 2026-08-18

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
