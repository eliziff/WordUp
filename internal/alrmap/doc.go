// Package alrmap is a Go port of the ALR executable code map: the offline
// Python 3.10 reference implementation (tools/alr-code-map, 149 named unit
// tests plus a 1,024-combination citation switch matrix) that maps the
// Alberta Law Review editing rules to concrete selectors, token operations,
// source bindings and noninterference checks. Nothing in it launches Word or
// edits a document in place.
//
// Layout mirrors the reference: kernels (spans, plans, masks, lexical and
// locator kernels), citations (complete note productions under ten
// switches), ooxml (byte-addressed XML tree, style inheritance, structural
// scan, reference-copy writer, title permission) and, in this package,
// references (note/source bindings), rules (finite serializers) and layout
// (unit conversion, style ownership, heading labels). PlanSourceTitles is
// the one function moved up from ooxml, because ooxml would otherwise import
// rules while layout imports ooxml.
//
// Offsets are Unicode code-point indices, as in Python, never byte offsets;
// nodes additionally keep their original UTF-8 byte boundaries. Refusals are
// returned errors carrying the reference messages; kernels.PermissionError
// stands for the one PermissionError and kernels.InvariantError for the
// AssertionError checks.
//
// Regular expressions are RE2. Every Python lookaround was replaced by a
// wider match plus a manual boundary check on the neighbouring code points:
// the DATE/PERCENT/WORDS/SPACES lexical patterns ((?<!\w), (?!\w),
// (?<![\w.,+\-]), (?![%\w]), \b, (?<=\S), (?=\S)), the OPAQUE naked-host
// lookbehind (?<![\w@]) (which additionally resumes one code point later so
// a host after a dot is still found), and the trailing citation lookaheads
// ((?= |,|$), (?= |$), (?=,| |$)), which became consuming groups whose
// extent is ignored. Python's Unicode \w, \d and \s are spelled out as
// \p{L}\p{N}_, \p{Nd} and the str.isspace set.
package alrmap
