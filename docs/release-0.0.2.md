# Release 0.0.2 readiness

OAKit `0.0.2` is a packaging-only hotfix for the immutable `0.0.1` npm
artifact. It changes no production source from the C3/R3, producer, rendering,
security, and 100% mutation evidence recorded for `0.0.1`.

## Required gates

- [x] Preserve the `oakit` bin mapping with npm 11-compatible path
      `dist/cli.js`.
- [x] Pack and install the artifact into a clean consumer and execute
      `node_modules/.bin/oakit --version`.
- [x] Retain the audited production source tree
      `4b7c749f58f908967b9868716bddd78e05a0fbcd`.
- [ ] Require all 11 CI jobs on the `0.0.2` release revision.
- [ ] Create tag and GitHub release `v0.0.2`.
- [ ] Complete provenance-enabled npm publishing and verify the public registry
      manifest retains `bin.oakit`.

## Parent evidence

- [`release-0.0.1.md`](release-0.0.1.md)
- [`evidence/0.0.1/release-gates.json`](evidence/0.0.1/release-gates.json)
