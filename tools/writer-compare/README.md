# Writer selection gates

Compare the current Go implementation with pinned pyOpenVBA on identical inputs.
Choose demonstrated capability first. If a more capable implementation is
materially slower, eliminate redundant work and port measured bottlenecks to Go.
Keep one production writer; alternative implementations may remain test oracles.

## Git acceptance

The editable workspace is the review/merge surface; DOTM is a built artifact.
Preserve existing workspace compatibility: text VBA, XML, structured form design,
separate assets and tests. Do not require agents to merge ZIP or binary VBA data.

- Repeated imports of identical bytes produce identical editable source.
- A no-op build does not modify tracked source, and preserves package content.
- Importing a no-op rebuild does not create source or form-design churn.
- A targeted edit produces only intended source/design changes; unknown data survives.
- Independent edits to separate modules, Ribbon and form design survive a Git merge
  and rebuild; overlapping edits remain visible conflicts, not silent last-writer wins.
- Test Git checkout line endings explicitly. Do not normalize exact XML expectations
  or binary assets merely to make comparisons pass.
- Run Git tests in disposable repositories; never alter the development worktree
  or include private specimens in public fixtures.

`compare.py` measures upstream unchanged saves, module edits, module add/rename/delete,
form reads, and control-caption/font edits. It records independent olefile stream hashes, package-part
changes and module/property checks, retaining local outputs. `go/` measures the
existing workspace build path; `--go-report` independently checks its outputs.
These paths have different overhead: do not attribute timings to language alone.
All operations compare form properties, designer text, structural control paths/order,
and hashes of parsed picture bytes. Structural paths retain unnamed nested records.
These parsed checks do not establish preservation of every unknown binary field;
the separate independent stream hashes expose changes for further investigation.
Run the oracle's defect controls with `python -m unittest discover -s tools/writer-compare`.

Git integration tests live in `internal/project/git_roundtrip_test.go` and include
real disposable branches, merges, explicit overlapping-source conflicts and autocrlf checkouts. Native caption/font inspection
uses `TestNativeWriterFormOutputs` with `WORDUP_NATIVE_TEST=1` and
`WORDUP_WRITER_REPORT` pointing to the local comparison report. Native compilation,
full form behavior and remaining writer capabilities are separate unfinished gates.

The native `open` operation accepts `named.disable_macros: true` on Windows.
This forces disabled macros for that open even in an execution-authorized host;
`macros_disabled_on_open` records the setting. It does not revoke the host's
authority to perform subsequent explicitly requested calls, nor is it a sandbox.
The form checks use indexed COM reads without running VBA or showing a form.
Font size is checked against the COM currency contract's exact scaled integer
string (12 points is `{"currency_scaled_10000":"120000"}`), not a rounded float.
