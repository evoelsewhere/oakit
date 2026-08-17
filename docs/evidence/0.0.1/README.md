# PowerPoint text-profile producer evidence

These manifests retain the producer observations used to certify source-free
text creation at `C3` and part-preserving text/transform editing at `R3` for
version `0.0.1`.

- Microsoft PowerPoint macOS `16.112` opened, reserialized, saved, and reopened
  both generated and edited packages without repair.
- LibreOffice Impress opened and resaved creation output, then performed two
  save/reopen cycles for edited output with PDF export and strict semantic
  verification.
- Google Slides performed controlled Drive API import/export for generated and
  edited packages. Both temporary presentations were deleted. Google normalized
  `45° + flipH + flipV` into the matrix-equivalent `-135°` rotation.
- Forced mutation artifacts from Reliability run
  [`32011779968`](https://github.com/evoelsewhere/oakit/actions/runs/32011779968)
  were canonicalized and passed the strict audit in run
  [`32017471364`](https://github.com/evoelsewhere/oakit/actions/runs/32017471364):
  16,190 killed, 4,159 compile errors, and zero missed across 88 files.
- Google Slides evidence was produced by run
  [`32023058314`](https://github.com/evoelsewhere/oakit/actions/runs/32023058314).

The matrix applies only to the profile IDs named in `producer-matrix.json`. It
is not a blanket claim for arbitrary PowerPoint features or edit operations.
