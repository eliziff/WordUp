# Reference bindings, layout ownership and evidence

Implemented reference functions and the VBA specimen are in the complete downloadable package attached to the associated conversation. This PR contains the implementation map and aggregate evidence, not the full runnable source tree. See the [publication note](../../tools/alr-code-map/README.md).

## 7. Source bindings survive numbering and ordering changes

`Note` stores a stable native identity, display label, numbering scope and complete parsed citation occurrences. `Binding` stores source identity, locator basis, ordered locator structure, qualifications and origin note separately. A native note ID is never assumed to be its displayed number.

`compile_graph` derives source keys from supported complete roots and exact parsed declarations. Editions and publication/source identifiers remain part of identity. Empty/malformed identities do not compare equal merely because both are unknown. Alias collisions produce unresolved results, not a first-match selection.

There are two deliberately different supra operations:

- `preserve_pointer(scope, label)` returns the unique note ID currently addressed by that label. It preserves an existing pointer without claiming the author chose the correct authority.
- `correct_alias_target(note_id, alias)` requires one exact declared source and a preceding full citation. It supplies a correction target only under that association; no surname, substring or fuzzy fallback exists.

A supported Ibid occurrence inherits its governing locator when its text omits one. That is distinct from converting a new source-wide citation into Ibid: such conversion must not silently inherit an earlier pinpoint. The reference kernel and graph retain source, edition/version, locator basis/value and qualifications rather than comparing only display text.

### Reordering example

```text
A: full declared source [Book], page 9
B: Ibid.  # bound to the same book and page 9
X is subsequently inserted between A and B, citing another source
```

`repair_ibid_after_reorder` takes the old graph and validated new note order/labels. It refuses a stale fingerprint or modified retained citation. It compares the new adjacency-derived meaning with B's old binding. If different, it renders a unique declared short form plus supra to A's new label and B's original inherited pinpoint. It does not reinterpret B as citing X.

Repair is restricted to a single bound Ibid citation in a note and an unambiguous preceding full source. Arbitrary within-note chains, missing old bindings and ambiguous multi-source predecessors remain unresolved. The output is a proposal string, not native field/revision mutation.

## 8. Typesetting: executable ownership and unit calculations

### Shared-style effect closure

`affected_styles` follows based-on descendants and linked partners. `style_uses` inspects actual paragraph/character/table references, including defaults, in supplied package parts. `require_style_ownership` refuses a style-definition change if an affected use lies outside the declared owned set.

This prevents a selection-only operation from redefining a shared style used by quotations elsewhere. Unknown references are not ignored. All relevant story parts must be supplied before claiming package-wide ownership: omitted headers, footers, text boxes or glossary parts are not certified.

This module does not itself change XML or native style definitions. Its effect analysis is a gate for a property writer, not a rendering result. Word's automatic style-redefinition control is also relevant but does not replace effect analysis. [W4]

### Explicit units and heading identities

`twips` requires a physical unit and decimal conversion. It distinguishes `0.2 in = 288 twips` from rounded `0.51 cm = 289 twips`. The guide's unitless footnote-blockquote value “1.5” is not converted by guessing.

`heading_labels` accepts stable heading IDs and levels, implements the mapped first three levels, checks parent order, resets subordinate counts, and produces labels under explicit numbered/unnumbered Introduction policy. It does not discover headings from short bold paragraphs, create native lists, rewrite document text or approximate rendered line counts.

The five-rendered-line blockquote rule, TOC/page convergence, linked headers, table geometry, native inheritance and PDF fidelity remain native-adapter/evidence work, not results established by these calculations.

## 9. Actual VBA source and its qualification boundary

The downloadable `native/ALRPageContract.bas` provides:

- `ALR_PageEnd`: finite VBA page-end contraction with positive/ascending validation.
- `ALR_TryPageDeletion`: complete explicit-page Ibid grammar returning the original snapshot's zero-based UTF-16 endpoint-prefix deletion.
- `ALR_ContractExplicitPageNote`: a bound `Footnote.Range` writer for only that deletion, with structural preflight and custom undo record.

The writer retains its duplicated note range, checks the complete snapshot and deletion text, and never performs a broad Find. It does not recreate note marks, change labels/dashes/styles, flatten fields or reconstruct notes. Existing revisions, fields, controls, hyperlinks, unexpected roles and structural boundaries refuse the generic path.

The Windows qualification specimen uses MSXML6 to inspect range XML. WordOpenXML is captured before starting the undo record, following this repository's recorded native observation. Previous tracking/screen state is restored on edit/error paths. This introduces no Python dependency into the `.dotm` design.

**The VBA was not compiled or run in Word for this delivery.** `ALR_ProbePageKernel` is probe source, not seven recorded native passes. MSXML qualification, native coordinates, Track Changes rejection, undo grouping, font preservation and save/reopen remain unverified. The specimen is not wired into the existing Ribbon or installed over the user's macro.

Range.Text replaces unformatted text; FormattedText can carry paragraph formatting when the paragraph mark is included; insertion can expand ranges/bookmarks. Document-level revision Count is not an all-stories count. These native properties are why reference-model tests do not substitute for Word qualification. [W5–W8]

## 10. Recorded evidence and reproducibility

The [aggregate evidence](../../tools/alr-code-map/evidence.json) records **149 passing named tests**, including the previous 53. Parameterized cases exercise token segmentation, structural wrappers and all 1,024 combinations of ten citation switches. These remain offline reference tests, not native Word passes.

| Corpus measurement | Count |
|---|---:|
| Manuscript footnotes parts inspected | 108 |
| Non-separator notes inspected | 17,261 |
| Single-paragraph notes | 17,122 |
| Complete supported syntax parses | 2,506 |
| Successful syntax/meaning round-trips | 2,506 |
| Already canonical supported notes | 2,261 |
| Supported notes with proposed differences | 245 |
| Proposed notes refused by structural/role gates | 33 |
| Note plans verified as XML-copy edits | 212 |
| Manuscripts with verified XML-copy edits | 36 |
| Changed text nodes | 213 |

No journal, written case-name, book-title or alias dictionaries were supplied to this corpus run, and bare at locators were not inferred as pages. Larger named grammars are exercised by synthetic fixtures. Multi-citation notes mean component counts may exceed note counts. These are coverage/noninterference measurements, not source-verification, correction-precision or native Word success rates.

The 33 refusals comprise 22 protected paragraph roles and 11 opaque structural dependencies. Earlier coarse locator-tail patches also refused 45 mixed-format notes. Smaller owned label, glyph, spacing and endpoint changes admitted those 45 without weakening structural gates.

The pipeline batches disjoint plans against one part snapshot, assembles bytes once and reparses/verifies once per part. It does not reparse a large footnotes part for every token. Weak parent links and a retained root avoid XML parent/child reference cycles delaying reclamation. The reported elapsed time is an offline Linux/Python measurement, not a macro-performance claim.

### Reproduce from the downloadable source package

```text
cd tools/alr-code-map
python -m unittest discover -v
python audit_corpus.py PATH/TO/alr-chatgpt-raw-xml.zip > corpus-summary.json
```

These commands apply to the full conversation download, not this documentation-only PR. No pip installation or network request is required. The audit changes only in-memory copies, not the source archive. Optional private-detail output contains manuscript text and must not be committed. This PR publishes no manuscript, template, raw bundle or exact private witness.

## 11. Scope not silently claimed

General Oxford-comma insertion, free-text semantic capitalization, honorific/entity inference, source-support verification, abstract writing, hierarchy/pinpoint translation, unmarked quote discovery, arbitrary bold-to-emphasis classification and complete rendered typesetting are not implemented by this package. Existing ALR manual tools remain retained requirements; their interfaces are unchanged and their native behaviour is not newly certified.

The noninterference claim is precise: for supported parsed structure and complete derived/supplied protections, the reference writer preserves all unowned XML structure/attributes/bytes and the expected text of every container. It is not a claim that passing regex tests makes every Word document safe.

## Sources

**Supplied policy:** ALR Style Guide 2025–2026, especially Primary Editing Guide, ALRisms, Pinpoints, Ibids/Supras/Internal References, Punctuation, Quotations, Numbers, Titles, Italics and Typesetting. Conflicts/unitless instructions remain explicit unresolved inputs. Personal quotation verification is not replaced by automatic green highlighting.

**Supplied implementation/corpus:** July 22 macro, its repository-reference snapshot, manuscript XML in the supplied raw archive, the prior ALR code map and 157-row crosswalk. Private inputs are not redistributed.

**Primary native references:**

- [W1: Range.SetRange](https://learn.microsoft.com/en-us/office/vba/api/word.range.setrange)
- [W2: Open XML Italic and toggle semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.italic)
- [W3: Open XML WebHidden](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.webhidden)
- [W4: Style.AutomaticallyUpdate](https://learn.microsoft.com/en-us/office/vba/api/word.style.automaticallyupdate)
- [W5: Range.Text](https://learn.microsoft.com/en-us/office/vba/api/word.range.text)
- [W6: Range.FormattedText](https://learn.microsoft.com/en-us/office/vba/api/word.range.formattedtext)
- [W7: Range.InsertBefore](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertbefore)
- [W8: Revisions](https://learn.microsoft.com/en-us/office/vba/api/word.revisions)
