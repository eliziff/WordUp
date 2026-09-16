# ALR editing rules: executable selectors and noninterference contracts

This expands the earlier ALR rule kernels into an **offline executable specification**. The complete tested source package is supplied in the associated conversation attachment; [the publication note](../tools/alr-code-map/README.md) explains why this PR contains the map and evidence rather than the full executable source. It is not a replacement `.dotm`, a Python dependency for recipients, or an assertion that Microsoft Word has executed these operations.

The inputs are the supplied ALR Style Guide 2025–2026, the July 22 macro, the manuscript XML, and the earlier adversarial code map. Journal instructions are identified as **policy**; parsers, restrictions and ownership checks are **implementation decisions**. A successful parse establishes supported syntax, not that an author cited the correct authority.

## Read the implementation map

1. [Protection extraction and exact XML writes](alr-code-map/01-protection-and-writes.md): actual field/revision/style state, original text-node addresses, batch mutation and whole-part noninterference checks.
2. [Citation grammars and executable style rules](alr-code-map/02-citations-and-rules.md): complete root productions, typed locators, independent switches, case-only title permissions and bounded value serializers.
3. [Reference bindings, layout ownership and evidence](alr-code-map/03-references-layout-and-evidence.md): stable note/source identities, Ibid repair after reordering, shared styles, VBA qualification source and actual corpus results.

## Entry points in the downloadable package

- `tools/alr-code-map/ooxml/`: byte-addressed tree, protective style inheritance, structural scan, writer and separate title permission.
- `tools/alr-code-map/citations/`: records/registry, lexical splitting, primary/secondary roots, typed locators and complete-note orchestration.
- `tools/alr-code-map/kernels/`: reusable span/plan model, source masks, lexical selectors, bounded locators and exact source/Perma utilities.
- `references.py`, `rules.py`, `layout.py`: concrete operations, not undefined classification helpers.

[Rule delta crosswalk](../tools/alr-code-map/rule-status.csv) · [Recorded offline evidence](../tools/alr-code-map/evidence.json)

The suite has **149 passing named tests**, plus parameterized switch/structural/run-splitting cases. The actual manuscript audit verifies **212 XML-copy note plans across 36 manuscripts**. **Native VBA compilation/execution, Track Changes rejection, Word save/reopen and rendered fidelity remain unverified.** No private manuscript or template is published here, and no macro has been deployed.
