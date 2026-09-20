# Citation grammars and executable style rules

Implemented functions described here are in the complete downloadable reference package. The [publication note](../../tools/alr-code-map/README.md) distinguishes that package from this documentation/evidence PR.

## 4. Full citation grammars and their owned tokens

**Policy:** the guide distinguishes citation syntax from explanatory footnote prose, paragraphs from pages, and source titles from quoted passages. The code uses those distinctions without assuming the whole footnote is citation syntax.

`normalize_note` scans supported balanced quotation, parenthesis and bracket syntax, then splits semicolons only at top level. Semicolons inside an angle-wrapped URI are not source separators. Every component must be fully consumed. An unexplained suffix rejects the whole parse.

| Implemented root | Required inputs and preserved identity |
|---|---|
| Ibid | Exact root; supported locator or none. Source resolution remains a separate graph operation. |
| Exact alias plus supra, or unnamed supra note N | Alias in the supplied exact symbol set, or genuinely unnamed pointer. Printed number remains a pointer until resolved. |
| Exact case name plus neutral/CanLII citation, or bare identifier | Exact case-name map when written name is present; enumerated court token; year and identifier preserved. |
| Reporter | Explicit year, exact reporter token, optional volume/series and starting page. The implemented bare-reporter grammar is not every named reporter citation. |
| Canadian legislation/regulation | Supported series/year/chapter or SOR/SI identifier. Written title needs an exact supplied boundary. Chapter/decision hyphens remain identity. |
| Journal | Authors, delimited title, year, volume/optional single or double issue, exact journal token, first page and optional locator. |
| Book | Exact supplied title boundary, supported optional edition, complete town/publisher/year tuple and optional locator. |
| Web | Authors, delimited title, complete named publication date, explicit online/type component and one complete wrapped HTTP(S) or Perma URI. |

Signals, short-form declarations and known qualification suffixes remain separate nodes. Recognizing them does not authorize reordering signals, inventing qualifications or changing source classification. Dictionaries are exact multi-maps: conflicts are refused. Canonical output spellings are recognized on a second run without duplicate identity entries; chained/nonterminal replacement tables are refused. Case-name rewriting is additionally limited to identity or removal of nonnumeric abbreviation periods, preventing a table from silently renaming a case.

### Speculative parser branches cannot own edits

An adversarial fixture made a journal author's name also match a registered case name. The draft parser emitted a case-name period deletion before establishing that a case identifier followed, then successfully parsed a journal citation. A similar failed statute-root match left a false title token.

The corrected `parse_primary` emits those edits and tokens only after the corresponding identifier production succeeds. Both counterexamples are retained. A partial name match cannot grant a root's edit permissions.

There is no general mixed-footnote prose parser, complete McGill parser, fuzzy authority resolver, hierarchy conversion or missing-metadata lookup.

### Numeric locators have different types

- `ParagraphLocator`: ordered integer/range tuple, full endpoints.
- `PageLocator`: ordered integer/range tuple, supported endpoint contraction.
- `LegalLocator`: unit, opaque identifiers and original separators.

The legal type preserves these distinctions:

```text
s 7        != art 7
ss 7 to 9  != ss 7, 9
s 3-2      is one identifier, not pages 3 through 2
```

The reference graph and serializer do not erase them before comparing Ibid meanings.

### Leaf patches, not replacement locator tails

For `Ibid at pp. 400 - 408.`, independent enabled operations change the page label, two delimiter-adjacent spaces, dash and redundant endpoint prefix. They do not delete/reinsert `400` merely to change its following dash. The complete parsed production supplies syntax evidence; the original token positions supply exact addresses.

Page contraction requires a valid ascending page range and the supported same-prefix/minimum-two-final-digits rule. The contracted endpoint must reconstruct the same integer. Paragraph endpoints are not shortened; legal identifiers never enter that function.

| Input | Canonical result |
|---|---|
| `Ibid at pp. 400-408.` | `Ibid at 400–08.` |
| `Ibid at paras 1847–1857.` | Unchanged |
| `Ibid, s. 3-2.` | `Ibid, s 3-2.` |

Bare `at 400–408` is not assigned a page basis from the digits. Explicit labels, a supported source production or supplied source-basis record establish that basis. After removing `pp.`, the accepted production's basis is retained for the round-trip check; a new parse of an unknown source cannot invent it.

**Ten independent mutation switches:** labels, range dashes, page contraction, pinpoint “and”, court order, journal names, case names, reporter names, dates and terminal punctuation. All 1,024 combinations are exercised on a supported composite fixture for parsed-meaning preservation and disabled behaviour. These are not 1,024 independent corpus/native Word tests.

CanLII court relocation owns only the existing parsed court/locator tail. Unknown parentheticals are not court nodes. A mixed-format move requiring token-preserving movement declines instead of homogenizing runs.

## 5. Title capitalization without disabling quote protection

**Policy:** source-title capitalization may change case; it does not permit changing the title's words. General spelling correction continues to protect the same text.

`plan_source_titles` derives a specific permission by fully parsing the enclosing note, checking XML structure and selecting only journal/book/web source-title interiors. Case and legislation titles are excluded from this general policy.

At commit, `title_guards` reparses the production and requires the TITLE-EN rule, casefold-equivalent old/new strings, and a complete read span inside a parsed title. Every field/revision/native guard and nested quotation/bracket/URI guard remains active. Only the outer guard exactly matching that title node is removed for the permitted case change.

A forged rule name cannot authorize spelling changes or edits to an author's name outside the title:

```text
Before:       “towards a judgement”
Case-only:    “Towards a Judgement”
Not allowed:  “Toward a Judgment”
```

`rules.title_case` implements a finite English small-word policy, first/after-colon/hyphen handling and exact-case exceptions. Unexplained acronyms/mixed-case tokens are retained and reported. Exception mappings must preserve casefolded identity. This is not a universal grammatical classifier. The guide's “with”/after-colon conflict is an explicit policy argument, not last-writer-wins. French and unknown-language interior capitalization are not guessed.

A title change invalidates reference graphs using its old text as an identity component. The package does not silently casefold source identities or migrate arbitrary aliases; the exact operation record is needed to rebuild/reassociate them.

## 6. Other executable rule families

| Family | Concrete operation | Negative boundary |
|---|---|---|
| Canadian spelling | Enumerated exact variant map, Unicode token boundaries, supported case preservation | No generated -our/-or or -ise/-ize rule; no quote/title/URI edits; grammatical pairs not inferred. |
| Dates | Complete named month/day/year validation, explicit Gregorian calendar, day-month-year rendering | No ambiguous numeric-date interpretation, invalid-date repair or date substitution inside titles. |
| Percent | Preserve exact integer/decimal digit string while replacing percent representation | Trailing zeros retained; malformed/URI/data/source contexts excluded. |
| Repeated spaces | Interword U+0020 runs inside authorized text | No broad whitespace matcher; no tabs, NBSP or paragraph/source cleanup. |
| Selected counts | Numerals one through ten to words; supplied larger positive integer grouped | No ordinary-number classifier or general words-to-digits parser. |
| Selected ordinal | Valid numeral/suffix, range 1–99 | Invalid suffixes, exponents and note markers cannot pass this role contract. |
| Selected ratio/count range | Exact two-endpoint grammar plus explicit role | No automatic classification of volume:issue, time or section identifiers. |
| Selected statutory reference | Complete identifier, supplied statutory identity and prose/citation context | Only labels change; hyphens/subdivisions remain. A bare paragraph letter lacks a complete section ID. |
| Quote separator | Supplied prose-quotation boundary and existing comma/period | No question/exclamation move, punctuation invention or automatic title/quote classification. |
| Quote qualifications | Explicit unique known flags, rendered without reordering | No emphasis-provenance inference; incompatible emphasis flags refused. |
| Lexical italics | Longest complete phrase intentions, including consuming roman exceptions | Roman intention cannot erase author/source emphasis without provenance or explicit selection. |
| Perma import | Exact CSV/whole-URI mapping | No substring replacement, path/query casefolding, capture-success assertion or native hyperlink mutation. |

A selected-role serializer is useful code but is not counted as unattended raw-manuscript recognition. Disabling a mutation switch never disables protection infrastructure required by other operations.

Next: [reference bindings, layout ownership and evidence](03-references-layout-and-evidence.md).
