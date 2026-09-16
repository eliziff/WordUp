# ALR executable code map

An isolated, standard-library **Python 3.10+ engineering reference** for mapping the supplied Alberta Law Review editing rules to concrete selectors, token operations, source bindings and noninterference checks. It does not add a Python runtime to WordUp or the recipient's `.dotm`, edit a live Word document, or change existing harness commands.

Read [the code map](../../docs/ALR-CODE-MAP.md).

## Run

```text
cd tools/alr-code-map
python -m unittest discover -v
python audit_corpus.py /path/to/alr-chatgpt-raw-xml.zip > corpus-summary.json
```

No pip installation or network access is required. Tests are synthetic and fast. The full corpus command is an explicit outer-loop audit, not part of routine test discovery. It reads the private input archive and verifies copies of eligible XML parts in memory; the original archive is unchanged.

## Modules

| File | Implemented boundary |
|---|---|
| `kernels.py` | Retained prior tested string kernels, quotation masks, patch validation, Perma mapping and bounded Ibid equivalence. |
| `ooxml.py` | Real XML/style protection extraction, original atom/byte addressing, operation-specific title permission, batch-copy writer with whole-part invariants. |
| `citations.py` | Whole-note supported citation productions, exact registry inputs, typed locators and leaf patches under ten independent switches. |
| `references.py` | Source/note/numbering separation, binding, pointer preservation, alias correction proposals and Ibid repair after reordering. |
| `rules.py` | Expanded exact lexical entries and selected title, number, statutory-reference and quote-boundary serializers. |
| `layout.py` | Explicit physical-unit conversion, shared-style effect/ownership analysis and supplied-heading label calculation. |
| `audit_corpus.py` | Aggregate-only private corpus audit; optional separately requested private witnesses. |
| `native/ALRPageContract.bas` | Actual narrow VBA parser/writer specimen; **not compiled, Word-tested, integrated or deployed**. |

`evidence.json` reports the executed tests and corpus counts without manuscript text. `rule-status.csv` accounts for all 157 previous registry entries and does not equate a partial helper with a complete native feature.

## Minimal parser example

```python
from citations import normalize_note

assert normalize_note('Ibid at pp. 400-408.').text == 'Ibid at 400–08.'
assert normalize_note('Ibid at paras 1847–1857.').text == 'Ibid at paras 1847–1857.'
assert normalize_note('Ibid, s. 3-2.').text == 'Ibid, s 3-2.'
```

A parser result is not a write authorization. Use `ooxml.plan_note` with real part/style bytes and `apply_batch` for the reference XML-copy checks. The writer rejects unsupported structural conditions; no `--unsafe` bypass is provided. Full-part snapshots and styles must remain unchanged between planning and applying.

## Privacy and evidence

Do not commit the private bundle, manuscripts, `.dotm` files, exact source witnesses, or `--private-details` output. The tool emits aggregates unless that option is explicitly supplied. Network archive creation is not performed.

String and XML tests are **not actual Word acceptance**. This delivery does not establish native Track Changes rejection, field insertion, undo grouping, save/reopen or rendered layout. The VBA probe source must never be described as already passed. Existing ALR Ribbon/manual features are not altered by this reference package.
