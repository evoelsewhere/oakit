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
- [x] Require all 11 CI jobs on the `0.0.2` release candidate; they passed in
      [run 32054867081](https://github.com/evoelsewhere/oakit/actions/runs/32054867081)
      on `5f6a9102bba8e483a023b28fc0eebee85461d92f`.
- [x] Require all 11 CI jobs on the final docs-only tag target; they passed in
      [run 32055043675](https://github.com/evoelsewhere/oakit/actions/runs/32055043675)
      on `a3eeb50ce7d78bb7bd92e53b32f41383aca1f57d`.
- [x] Create tag and [GitHub release `v0.0.2`](https://github.com/evoelsewhere/oakit/releases/tag/v0.0.2).
- [x] Complete provenance-enabled npm publishing in
      [run 32055249795](https://github.com/evoelsewhere/oakit/actions/runs/32055249795)
      and verify the public registry manifest retains
      `bin.oakit = dist/cli.js`.
- [x] Install `@evoelsewhere/oakit@0.0.2` from the public registry into a clean
      consumer, run `oakit 0.0.2` through its installed bin, and import the ESM
      API.

## Parent evidence

- [`release-0.0.1.md`](release-0.0.1.md)
- [`evidence/0.0.1/release-gates.json`](evidence/0.0.1/release-gates.json)
- [`evidence/0.0.2/release-gates.json`](evidence/0.0.2/release-gates.json)
