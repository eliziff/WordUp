# Protection extraction and exact writes

The functions below are implemented in the complete downloadable reference package supplied with the associated conversation. This PR publishes their implementation map and aggregate evidence; see the [publication limitation](../../tools/alr-code-map/README.md). They are not claims of native Word execution.

## 1. What is now code instead of a missing helper

| Former prerequisite | Implemented entry | Deliberate boundary |
|---|---|---|
| Protect fields and existing revisions | `ooxml.scan`: field stack, revision wrappers, enclosing structural ranges | Actual XML inputs, not supplied boolean flags. Unknown/malformed structures block mutation. |
| Respect inherited italics | `Styles.chain`, `Styles.run`, `Styles.paragraph_risk` | Selected protective properties with actual inheritance/toggle semantics; not a universal font/layout resolver. |
| Edit the occurrence inspected | `bind`, `prepare`, `apply_batch` | Snapshot-bound container, text atom and original byte range. Never re-find text in a larger range. |
| Preserve runs | `locator_edits` and the text-node writer | Separate label, dash, spacing and endpoint deletions. Retained digits keep their original runs. |
| Recognize a citation | `citations.normalize_note` | Complete bounded productions, top-level separators, explicit dictionaries, typed locators, no suffix salvage. |
| Do not edit title words | Parsed `source-title` tokens plus quotation/URI/native guards | Title capitalization is a separate re-derived permission; it does not authorize spelling correction. |
| Keep supras and ibids correct | `references.compile_graph`, `preserve_pointer`, `repair_ibid_after_reorder` | Note IDs/labels are inputs; source, version, locator and qualification remain separate. No Word field insertion is simulated. |
| Changing a style is local | `layout.require_style_ownership` | Based-on descendants and linked partners checked against supplied story usages; omitted parts cannot be certified. |
| Apply ALR title/number rules | `rules.title_case`, `selected_number`, `selected_legal_reference` | Finite selected-role serializers, not hidden title/part-of-speech/referent classifiers. |

The 157-entry delta crosswalk records partial implementations as partial. One implemented helper does not make an entire native feature complete.

## 2. Actual XML to proposed edit to verified XML copy

The package's `plan_note` takes a `Scan`, an XML container ID and a `Registry`. `Scan` comes from actual package part bytes plus `styles.xml`, not concatenated paragraph strings. `apply_batch` returns new XML bytes and the exact original byte slices changed. Nothing is written in place.

`footnote:17` means the source XML note ID, not displayed footnote 17. Every plan in a batch refers to the same original part/styles snapshot.

### Separate address systems

The XML projection uses Python code-point positions and retains an `Atom(start, end, node, signature)` for every editable `w:t`. Each node retains its original UTF-8 byte boundaries. The VBA specimen instead uses the exact `Range.Text` snapshot and native story-relative UTF-16 positions. `utf16_offset` converts a string offset; it does not turn flattened XML into a native Word range.

Word's `SetRange` counts story-relative hidden and nonprinting characters. A native adapter must retain the actual note/story range; a `Document.Range` constructed from footnote-local offsets is not equivalent. [W1 in the sources chapter]

### Exact write algorithm

`prepare(snapshot, bound)` performs all of the following before returning a mutation:

1. Compare the complete part/styles/part-name fingerprint and container text digest.
2. Re-derive mandatory guards from the scanner. Supplying an empty `Plan.protected` cannot bypass them.
3. Validate each edit's old text, read footprint, write footprint, authority and overlap constraints. Coincident insertion boundaries are conflicts too.
4. Map the original occurrence into its original text nodes. Never relocate it with `Find`.
5. Permit pure deletions across ordinary run boundaries while preserving surviving runs. Insertions/replacements need one compatible property signature over the leaf read span; otherwise decline.
6. Refuse structural gaps, empty/self-closing text-node rewrites, and changes requiring an invented `xml:space` attribute.

`apply_batch` collects node-specific replacements from disjoint container plans; assembles output once from unchanged original byte slices and escaped new text; and parses/rescans the output once. It then checks every container against its expected text, including untouched containers, and compares the complete element sequence, tags, attributes, child counts and all unowned direct text.

Only `w:t` character-data regions are writable. Comments or processing instructions inside a text node cause refusal rather than accidental deletion. XML comments, relationship references, field instructions, run properties, note marks and bytes outside the owned regions are retained exactly. This is stronger than visible-text comparison and narrower than claiming that a subsequent Word save preserves identical bytes.

The proof is about this representation and writer: unchanged slices are copied unchanged; tree comparison detects unowned structural changes; protected read/write intersections cannot be committed. It does not establish that unmarked quotations are discoverable or that Word preserves revision grouping in a future native adapter.

## 3. Deriving protections from the document

### Structural state persists across paragraphs

The main body is one logical container. Each non-separator footnote/endnote is an independent container. Paragraph boundaries do not reset quotation or field state.

| XML evidence | State transition or mask |
|---|---|
| `fldChar begin/separate/end` | Push/update/pop a nested field stack; protect the full interval and boundary events. |
| `instrText` | Never project as editable visible text. An orphan instruction quarantines the container. |
| `ins`, `del`, `moveFrom`, `moveTo`, `delText` | Opaque revision interval, not flattened current prose. |
| Bookmark/comment/permission/move range boundaries | Pair IDs across paragraphs; protect the complete interval and endpoint events. |
| Unmatched, duplicate or unsupported range pair | Quarantine the logical container rather than one paragraph. |
| Hyperlink, content control, custom XML, table, drawing, equation, unknown wrapper | Opaque enclosing interval, including nested text resembling eligible tokens. |
| Native note-reference mark | Separate structural event, never an ordinary digit. A leading note mark is not copied into citation text. |
| Tab, break, body footnote-reference mark, symbol | Explicit structural atom; no generic whitespace consumption. |
| `pPrChange` or `rPrChange` | Protect the affected paragraph rather than ignoring the formatting revision. |

A digit-only field result remains a field. Inspecting only a candidate's `Fields.Count` misses a field whose root encloses it; the scanner processes enclosing structure first.

### Inheritance is executed

`Styles` reads document defaults, default paragraph/character styles, explicit style references and every `basedOn` ancestor. Missing styles, type mismatches and cycles are failures, not evidence of ordinary formatting.

For the mapped toggle properties, a true style value flips inherited state; false leaves it unchanged. A direct run value sets it absolutely:

- Parent italic true plus child style italic true: false.
- Parent italic true plus child style italic false: true.
- Parent italic true plus direct italic false: false.

This follows documented toggle semantics for the explicitly enumerated properties. `webHidden` and `dstrike` use ordinary absolute on/off handling, not that toggle logic. Adversarial fixtures cover the distinction. [W2–W3]

Resolved italics, bold, caps, underline, vertical placement and non-English language can produce lexical protections. Hidden/web-hidden text is a hard protection. Inherited paragraph indentation, numbering, outline roles, frames and known quote-style names also veto ordinary lexical operations. Unsupported table, numbering and script-dependent formatting are not approximated into plain text.

The insertion property signature is deliberately stricter than visual equality: it includes the styles snapshot and relevant raw paragraph/run property trees. Two runs that look alike may still be refused. This is not a complete effective-style engine.

### Lexical masks and authorization

The retained stack scanner protects balanced nested quotation/title spans, carries state across paragraphs and supports repeated opening quotes at quoted-paragraph beginnings. Unsupported, malformed or ambiguous delimiters quarantine the logical container. Multiple ambiguous ASCII quote pairs remain conservatively blocked.

URI/email and bracket masks are independent. Broad URL masking may include adjacent punctuation to avoid damage; it is not the exact Perma replacement selector. Authorizing a body region does not disable any nested protections.

Identical unmarked quoted and authored strings remain indistinguishable from these inputs. Normal style and absence of quotation marks do not establish authorial-text authority. `plan_author_text` remains advisory until an explicit reviewed scope/import origin is supplied.

Next: [citation grammars and executable style rules](02-citations-and-rules.md).
