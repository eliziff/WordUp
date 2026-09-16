# ALR code-map publication and evidence

This directory accompanies [the expanded implementation map](../../docs/ALR-CODE-MAP.md).

**Publication limitation:** the complete executable reference package was built and tested in the associated conversation, but repeated source-file upload requests were blocked by the connector with an indeterminate safety-status error. This PR therefore publishes documentation, the rule crosswalk and aggregate evidence, not a complete runnable source tree. Incomplete source files were removed rather than leaving broken imports or claiming that the code had been published.

The downloadable `ALR_Executable_Code_Map_Expansion.zip` attached to that conversation contains the complete standard-library Python reference implementation, 149 tests, corpus-audit command and unexecuted VBA qualification specimen. The package uses the source paths documented in the implementation map. It adds no Python dependency to the recipient's `.dotm` and is not a deployed replacement macro.

`evidence.json` distinguishes executed offline tests, actual manuscript XML inspection, verified XML-copy edits and native Word checks that have **not** been performed. `rule-status.csv` maps all 157 previous feature IDs; blank entries mean not implemented in the reference expansion, and partial/selected statuses must not be read as complete native features.

No private manuscripts, raw source bundle, templates or exact corpus witnesses are published here. No existing WordUp command or ALR Ribbon action is changed.
