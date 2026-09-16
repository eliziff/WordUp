"""Infer a typesetting and citation profile for every olj-db journal with
2025/2026 output from the published articles' final-contract packages
(per-page line records with types, sizes and boxes, plus the native PDF
font summary). No PDF is opened here; the packages are read-only inputs.

Each profile field carries the measurement, the number of articles that agree,
a confidence in [0, 1] and an evidence list (article id, page, value). Fields
below the confidence floor are emitted with value null so a generated macro
skips that rule instead of guessing.

Usage:
    python tools/journal-profiles/infer_profiles.py --out internal/journal/profiles
"""
from __future__ import annotations

import argparse
import collections
import json
import os
import re
import sqlite3
import statistics

OAJD_ROOT = r"C:\Users\elias\Desktop\Open Access Journals Database"
DB_PATH = os.path.join(OAJD_ROOT, "oajd", "journals.db")
YEARS = ("2025", "2026")
PX_PER_PT = 2.0  # contract boxes are rendered at zoom 2 (144 dpi): 6 x 9 in pages are 864 x 1296 px
CONFIDENCE_FLOOR = 0.6
REPOSITORY_FURNITURE = re.compile(r"published by|downloaded from|digitalcommons|commons\.|scholarship|https?://|doi\.org|cambridge core|erudit|persée|©", re.I)

# ---------------------------------------------------------------- helpers ---


def pct(values, q):
    if not values:
        return None
    s = sorted(values)
    k = max(0, min(len(s) - 1, int(round(q * (len(s) - 1)))))
    return s[k]


def mode(values):
    values = [v for v in values if v is not None]
    if not values:
        return None
    return collections.Counter(values).most_common(1)[0][0]


def agreement(values):
    """(winning value, share of articles agreeing, count)."""
    values = [v for v in values if v is not None]
    if not values:
        return None, 0.0, 0
    counter = collections.Counter(values)
    value, count = counter.most_common(1)[0]
    return value, count / len(values), len(values)


def field(values, evidence, rounding=None, floor=CONFIDENCE_FLOOR, minimum=2):
    value, share, n = agreement(values)
    out = {"value": None, "confidence": round(share, 2), "articles": n, "evidence": evidence[:6]}
    if value is None or share < floor or n < minimum:
        return out
    if rounding is not None and isinstance(value, (int, float)):
        value = round(value, rounding)
    out["value"] = value
    return out


def quantize(value, step):
    if value is None:
        return None
    return round(round(value / step) * step, 2)


def px_to_pt(v):
    return None if v is None else v / PX_PER_PT


def px_to_in(v):
    return None if v is None else v / PX_PER_PT / 72.0


def ev(art, page, value):
    return {"article_id": art["article_id"], "page": page, "value": value}


def alignment(x0, x1, width, tolerance=0.05):
    if x0 is None or x1 is None or not width:
        return None
    center = (x0 + x1) / 2
    return "center" if abs(center - width / 2) < width * tolerance else "left"


def text_case(text):
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return None
    upper = sum(1 for c in letters if c.isupper())
    if upper / len(letters) > 0.9:
        return "upper"
    words = [w for w in re.split(r"\s+", text) if w and w[0].isalpha()]
    if not words:
        return None
    cap = sum(1 for w in words if w[0].isupper())
    return "title" if cap / len(words) >= 0.75 else "sentence"


# --------------------------------------------------------- article parsing ---


def load_pages(path):
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                yield json.loads(line)


CIT_PATTERNS = {
    "ibid_capitalized": re.compile(r"(?<![\w])Ibid(?![\w])"),
    "ibid_lower": re.compile(r"(?<![\w])ibid(?![\w])"),
    "ibid_period": re.compile(r"(?<![\w])[Ii]bid\."),
    "ibid_comma": re.compile(r"(?<![\w])[Ii]bid,"),
    "ibid_at": re.compile(r"(?<![\w])[Ii]bid at\b"),
    "supra_note": re.compile(r"(?<![\w])supra note \d+"),
    "supra_comma_note": re.compile(r"(?<![\w])supra, note \d+"),
    "at_para": re.compile(r"\bat paras? \d"),
    "at_para_period": re.compile(r"\bat paras?\. \d"),
    "at_pilcrow": re.compile(r"\bat ¶ ?\d"),
    "at_page": re.compile(r"\bat \d+(?![\w:])"),
    "at_p": re.compile(r"\bat pp?\.? \d"),
    "range_endash": re.compile(r"\d–\d"),
    "range_hyphen": re.compile(r"\bat \d+-\d+"),
    "emphasis_brackets": re.compile(r"\[emphasis (?:added|in original)\]"),
    "emphasis_parens": re.compile(r"\(emphasis (?:added|in original)\)"),
    "section_s_period": re.compile(r"(?<![\w])ss?\. \d"),
    "section_s_bare": re.compile(r"(?<![\w])ss? \d"),
    "online_colon": re.compile(r"\bonline: ?<?https?://", re.I),
    "online_angle": re.compile(r"\bonline: ?<https?://", re.I),
    "perma": re.compile(r"perma\.cc/", re.I),
    "et_al_period": re.compile(r"\bet al\."),
    "et_al_bare": re.compile(r"\bet al(?![\w.])"),
    "neutral_citation": re.compile(r"\b(?:19|20)\d\d [A-Z]{2,6} \d+"),
    "canlii": re.compile(r"\bCanLII\b"),
    "eg_bare": re.compile(r"(?<![\w])eg(?![\w.])"),
    "eg_period": re.compile(r"(?<![\w])e\.g\."),
    "ie_bare": re.compile(r"(?<![\w])ie(?![\w.])"),
    "ie_period": re.compile(r"(?<![\w])i\.e\."),
    "curly_double": re.compile(r"[\u201c\u201d]"),
    "straight_double": re.compile(r"\""),
    "see_signal_upper": re.compile(r"(?<![\w])See(?: also| generally)?\b"),
    "see_signal_lower": re.compile(r"(?<![\w])see(?: also| generally)?\b"),
}

HEADING_NUMBERING = [
    ("part_roman", re.compile(r"^(?:PART|Part) [IVXLC]+\b")),
    ("roman_period", re.compile(r"^[IVXLC]+\.\s")),
    ("roman_bare", re.compile(r"^[IVXLC]+\s+[A-Z\u201c]")),
    ("arabic_multi", re.compile(r"^\d+\.\d+")),
    ("arabic_period", re.compile(r"^\d+\.\s")),
    ("arabic_paren", re.compile(r"^\(?\d+\)\s")),
    ("alpha_period", re.compile(r"^[A-Z]\.\s")),
    ("alpha_paren", re.compile(r"^\(?[a-zA-Z]\)\s")),
    ("unnumbered", re.compile(r"^[A-Za-z\u201c\"]")),
]
NUMBERING_PREFIX = re.compile(r"^(?:(?:PART|Part) [IVXLC]+[.:]?|[IVXLC]+\.|\d+(?:\.\d+)*\.?|\(?\d+\)|[A-Z]\.|\(?[a-zA-Z]\))\s+")
NOTE_NUMBER = [
    ("number_gap", re.compile(r"^\d+\s{2,}\S")),
    ("number_period", re.compile(r"^\d+\.\s")),
    ("number_space", re.compile(r"^\d+\s\S")),
    ("symbol", re.compile(r"^[†‡*§¶]")),
]


def analyze_article(dataset, article_id, contract_dir, doc_type):
    pages_path = os.path.join(contract_dir, "pages.jsonl")
    summary_path = os.path.join(contract_dir, "digitalborn_native_summary.json")
    if not os.path.exists(pages_path):
        return None
    summary = {}
    if os.path.exists(summary_path):
        with open(summary_path, encoding="utf-8") as f:
            summary = json.load(f)
    pages = list(load_pages(pages_path))
    if not pages:
        return None
    role_body = summary.get("font_role_body") or {}
    role_note = summary.get("font_role_note") or {}
    art = {
        "article_id": article_id, "dataset": dataset, "doc_type": doc_type, "pages": len(pages),
        "page_size_px": None,
        "body_font": role_body.get("font"), "body_size": role_body.get("size"),
        "note_font": role_note.get("font"), "note_size": role_note.get("size"),
        "left_by_parity": {0: [], 1: []}, "right_by_parity": {0: [], 1: []}, "top": [], "bottom": [],
        "line_heights": [], "note_x0": [], "note_lines": [], "note_number_styles": [],
        "headers": [], "numbers": [], "headings": [], "block_quote_indent": [], "block_quote_size": [],
        "toc_pages": 0, "doc_title": [], "authors": [], "first_page": None, "first_page_author_note": [],
        "text_widths": [],
    }
    # Locate the article's first page: the first page carrying doc_title lines,
    # else the first page with body text (repository cover pages precede it).
    for page in pages:
        if any((ln.get("type") or reg.get("type")) == "doc_title" for reg in page.get("regions") or [] for ln in reg.get("lines") or []):
            art["first_page"] = page.get("pdf_page")
            break
    if art["first_page"] is None:
        for page in pages:
            if any((ln.get("type") or reg.get("type")) == "text" for reg in page.get("regions") or [] for ln in reg.get("lines") or []):
                art["first_page"] = page.get("pdf_page")
                break
    for page in pages:
        pn = page.get("pdf_page")
        ps = (page.get("region_payload") or {}).get("page_size")
        if ps and len(ps) == 2 and art["page_size_px"] is None:
            art["page_size_px"] = [ps[0], ps[1]]
        width = ps[0] if ps and len(ps) == 2 else None
        height = ps[1] if ps and len(ps) == 2 else None
        if art["first_page"] is None or pn is None or pn < art["first_page"]:
            continue
        is_first = pn == art["first_page"]
        interior = pn > art["first_page"]
        body_lines, note_y0s, footer_ys = [], [], []
        title_seen, body_seen = False, False
        for reg in page.get("regions") or []:
            rtype = reg.get("type")
            if rtype == "table_of_contents":
                art["toc_pages"] += 1
            for ln in reg.get("lines") or []:
                ltype = ln.get("type") or rtype
                bbox = ln.get("bbox") or {}
                text = (ln.get("text") or "").strip()
                size = ln.get("native_pdf_median_font_size")
                x0, x1, y0, y1 = bbox.get("x0"), bbox.get("x1"), bbox.get("y0"), bbox.get("y1")
                if ltype == "doc_title" and text:
                    title_seen = True
                    art["doc_title"].append({"page": pn, "text": text, "size": size, "x0": x0, "x1": x1, "width": width})
                elif ltype == "text":
                    body_seen = True
                    if interior and x0 is not None and size and art["body_size"] and abs(size - art["body_size"]) < 0.6:
                        body_lines.append((x0, x1, y0, y1))
                elif ltype == "footnote":
                    if text:
                        art["note_lines"].append(text)
                        kind = next((name for name, rx in NOTE_NUMBER if rx.search(text)), None)
                        if kind:
                            art["note_number_styles"].append(kind)
                        if is_first and re.match(r"^[†‡*§¶]|^\D", text):
                            art["first_page_author_note"].append(text)
                    if x0 is not None:
                        art["note_x0"].append(x0)
                    if y0 is not None:
                        note_y0s.append(y0)
                    if y1 is not None and interior:
                        art["bottom"].append((height - y1) if height else None)
                elif ltype == "header" and text and interior:
                    art["headers"].append({"page": pn, "text": text, "x0": x0, "x1": x1, "y0": y0, "width": width})
                elif ltype == "footer" and y0 is not None:
                    if REPOSITORY_FURNITURE.search(text):
                        footer_ys.append(y0)
                elif ltype == "number" and text and interior:
                    art["numbers"].append({"page": pn, "text": text, "x0": x0, "x1": x1, "y0": y0, "y1": y1, "width": width, "height": height})
                elif ltype in ("paragraph_title", "byline") and text:
                    if is_first and not body_seen:
                        # Lines between the title and the first body paragraph
                        # on the article's first page name the authors (or a
                        # section banner above the title).
                        if title_seen:
                            art["authors"].append({"page": pn, "text": text, "size": size, "x0": x0, "x1": x1, "width": width})
                        continue
                    if "....." in text or ltype == "byline":
                        continue
                    art["headings"].append({"page": pn, "text": text, "size": size, "x0": x0, "x1": x1, "width": width})
                elif ltype == "block_quote" and x0 is not None:
                    art["block_quote_indent"].append(x0)
                    if x1 is not None:
                        art["text_widths"].append(x1 - x0)
                    if size:
                        art["block_quote_size"].append(size)
        # Page numbers printed on the same baseline as repository furniture belong to the repository.
        if footer_ys:
            art["numbers"] = [n for n in art["numbers"] if not (n["page"] == pn and n.get("y0") is not None and any(abs(n["y0"] - fy) < 6 for fy in footer_ys))]
        if body_lines and width:
            xs0 = [l[0] for l in body_lines]
            xs1 = [l[1] for l in body_lines if l[1] is not None]
            ys0 = [l[2] for l in body_lines if l[2] is not None]
            parity = pn % 2
            art["left_by_parity"][parity].append(pct(xs0, 0.1))
            if xs1:
                art["right_by_parity"][parity].append(width - pct(xs1, 0.9))
            if ys0:
                art["top"].append(min(ys0))
            heights = sorted(l[3] - l[2] for l in body_lines if l[2] is not None and l[3] is not None)
            if len(heights) >= 4:
                art["line_heights"].extend(heights[len(heights) // 4: -(len(heights) // 4)])
            if not note_y0s and height:
                ys1 = [l[3] for l in body_lines if l[3] is not None]
                if ys1:
                    art["bottom"].append(height - max(ys1))
    art["bottom"] = [b for b in art["bottom"] if b is not None]
    return art


# ---------------------------------------------------------- aggregation ----


def classify_header(text, journal_name, journal_abbrev, title_words, author_words):
    t = text.lower()
    classes = set()
    if journal_name and journal_name.lower() in t:
        classes.add("journal_name")
    elif journal_abbrev and journal_abbrev.lower().replace(".", "") in t.replace(".", ""):
        classes.add("journal_abbrev")
    if re.search(r"\bvol(?:ume)?\.?\s?\d+", t) or re.search(r"\b(?:19|20)\d\d\b", t):
        classes.add("volume_or_year")
    words = set(re.findall(r"[a-z]{4,}", t))
    if title_words and len(words & title_words) >= max(2, len(title_words) // 3):
        classes.add("article_title")
    if author_words and words & author_words:
        classes.add("author")
    if not classes:
        classes.add("other")
    return classes


def level_scheme(articles):
    """Numbering pattern per heading level, derived from the observed pattern
    hierarchy: roman/part above alpha above arabic above unnumbered."""
    order = ["part_roman", "roman_period", "roman_bare", "arabic_period", "arabic_paren", "arabic_multi", "alpha_period", "alpha_paren", "unnumbered"]
    per_article = []
    evid = []
    for a in articles:
        counts = collections.Counter()
        for h in a["headings"]:
            kind = next((name for name, rx in HEADING_NUMBERING if rx.search(h["text"])), None)
            if kind:
                counts[kind] += 1
        if not counts:
            continue
        present = [k for k in order if counts.get(k, 0) >= 1]
        per_article.append(tuple(present[:3]))
        evid.append(ev(a, a["headings"][0]["page"], [h["text"][:50] for h in a["headings"][:3]]))
    value = field(per_article, evid, floor=0.5)
    if value["value"] is not None:
        value["value"] = list(value["value"])
    return value


def aggregate(dataset, registry, articles):
    prof = {"dataset": dataset, "journal": registry.get("name"), "abbrev": registry.get("abbrev"),
            "platform": registry.get("platform"), "years": list(YEARS),
            "articles_measured": [a["article_id"] for a in articles], "units": "points and inches; boxes measured at 144 dpi"}
    if not articles:
        prof["status"] = "no_contracts"
        return prof
    # Page size
    sizes, evid = [], []
    for a in articles:
        if a["page_size_px"]:
            key = (quantize(px_to_pt(a["page_size_px"][0]), 1), quantize(px_to_pt(a["page_size_px"][1]), 1))
            sizes.append(key)
            evid.append(ev(a, a["first_page"], {"width_pt": key[0], "height_pt": key[1]}))
    value, share, n = agreement(sizes)
    prof["page"] = {"width_pt": value[0] if value else None, "height_pt": value[1] if value else None,
                    "width_in": round(value[0] / 72, 2) if value else None, "height_in": round(value[1] / 72, 2) if value else None,
                    "confidence": round(share, 2), "articles": n, "evidence": evid[:6]}

    # Margins: inner/outer from page parity; symmetric when parity does not matter.
    def parity_margin(key):
        odd, even, evid = [], [], []
        for a in articles:
            o = a[key][1]
            e = a[key][0]
            if o:
                odd.append(quantize(px_to_in(statistics.median(o)), 0.05))
            if e:
                even.append(quantize(px_to_in(statistics.median(e)), 0.05))
            if o or e:
                evid.append(ev(a, None, {"odd": quantize(px_to_in(statistics.median(o)), 0.05) if o else None, "even": quantize(px_to_in(statistics.median(e)), 0.05) if e else None}))
        return field(odd, evid, rounding=2, floor=0.5), field(even, evid, rounding=2, floor=0.5)
    left_odd, left_even = parity_margin("left_by_parity")
    right_odd, right_even = parity_margin("right_by_parity")
    mirrored = None
    if left_odd["value"] is not None and left_even["value"] is not None and right_odd["value"] is not None and right_even["value"] is not None:
        mirrored = abs(left_odd["value"] - left_even["value"]) >= 0.15 and abs(left_odd["value"] - right_even["value"]) < 0.15
    prof["margins_in"] = {
        "mirrored": mirrored,
        "left_odd_pages": left_odd, "left_even_pages": left_even,
        "right_odd_pages": right_odd, "right_even_pages": right_even,
        "top": field([quantize(px_to_in(statistics.median(a["top"])), 0.05) for a in articles if a["top"]],
                     [ev(a, None, quantize(px_to_in(statistics.median(a["top"])), 0.05)) for a in articles if a["top"]], rounding=2, floor=0.5),
        "bottom": field([quantize(px_to_in(statistics.median(a["bottom"])), 0.05) for a in articles if a["bottom"]],
                        [ev(a, None, quantize(px_to_in(statistics.median(a["bottom"])), 0.05)) for a in articles if a["bottom"]], rounding=2, floor=0.5),
    }
    prof["body"] = {
        "font": field([a["body_font"] for a in articles], [ev(a, None, a["body_font"]) for a in articles]),
        "size_pt": field([quantize(a["body_size"], 0.5) for a in articles], [ev(a, None, a["body_size"]) for a in articles], rounding=1),
        "line_height_pt": field([quantize(px_to_pt(statistics.median(a["line_heights"])), 0.5) for a in articles if a["line_heights"]],
                                [ev(a, None, quantize(px_to_pt(statistics.median(a["line_heights"])), 0.5)) for a in articles if a["line_heights"]], rounding=1, floor=0.5),
    }
    note_indents, note_ev = [], []
    for a in articles:
        lefts = a["left_by_parity"][0] + a["left_by_parity"][1]
        if a["note_x0"] and lefts:
            v = quantize(px_to_in(pct(a["note_x0"], 0.1) - statistics.median(lefts)), 0.05)
            note_indents.append(v)
            note_ev.append(ev(a, None, v))
    prof["footnotes"] = {
        "font": field([a["note_font"] for a in articles], [ev(a, None, a["note_font"]) for a in articles]),
        "size_pt": field([quantize(a["note_size"], 0.5) for a in articles], [ev(a, None, a["note_size"]) for a in articles], rounding=1),
        "present": field([len(a["note_lines"]) > 0 for a in articles], [ev(a, None, len(a["note_lines"])) for a in articles]),
        "number_style": field([mode(a["note_number_styles"]) for a in articles if a["note_number_styles"]],
                              [ev(a, None, a["note_lines"][:1]) for a in articles if a["note_number_styles"]], floor=0.5),
        "indent_vs_body_in": field(note_indents, note_ev, rounding=2, floor=0.5),
        "first_page_author_note": field([len(a["first_page_author_note"]) > 0 for a in articles],
                                        [ev(a, a["first_page"], a["first_page_author_note"][:1]) for a in articles], floor=0.5),
    }
    # Headings
    cases, sizes_h, aligns, h_ev = [], [], [], []
    for a in articles:
        if not a["headings"]:
            continue
        top_size = max((h["size"] or 0) for h in a["headings"])
        top = [h for h in a["headings"] if h["size"] and abs(h["size"] - top_size) < 0.3]
        cs = [text_case(NUMBERING_PREFIX.sub("", h["text"])) for h in top]
        al = [alignment(h["x0"], h["x1"], h["width"]) for h in top]
        cases.append(mode(cs))
        sizes_h.append(quantize(top_size, 0.5))
        aligns.append(mode(al))
        h_ev.append(ev(a, top[0]["page"], top[0]["text"][:60]))
    prof["headings"] = {
        "numbering_levels": level_scheme(articles),
        "top_level_case": field(cases, h_ev, floor=0.5),
        "top_level_size_pt": field(sizes_h, h_ev, rounding=1, floor=0.5),
        "top_level_alignment": field(aligns, h_ev, floor=0.5),
        "samples": [h["text"][:80] for a in articles[:3] for h in a["headings"][:4]],
    }
    # Block quotes
    bq_vals, bq_ev = [], []
    for a in articles:
        lefts = a["left_by_parity"][0] + a["left_by_parity"][1]
        if a["block_quote_indent"] and lefts:
            v = quantize(px_to_in(pct(a["block_quote_indent"], 0.2) - statistics.median(lefts)), 0.05)
            bq_vals.append(v)
            bq_ev.append(ev(a, None, v))
    prof["block_quotes"] = {
        "indent_in": field(bq_vals, bq_ev, rounding=2, floor=0.5),
        "size_pt": field([quantize(mode(a["block_quote_size"]), 0.5) for a in articles if a["block_quote_size"]],
                         [ev(a, None, mode(a["block_quote_size"])) for a in articles if a["block_quote_size"]], rounding=1, floor=0.5),
        "articles_with_block_quotes": sum(1 for a in articles if a["block_quote_indent"]),
    }
    # Running heads and page numbers
    name = registry.get("name") or ""
    abbrev = registry.get("abbrev") or ""

    def head_summary(parity):
        kinds, evid = [], []
        for a in articles:
            title_words = set(re.findall(r"[a-z]{4,}", " ".join(t["text"] for t in a["doc_title"]).lower()))
            author_words = set(re.findall(r"[a-z]{3,}", " ".join(t["text"] for t in a["authors"]).lower()))
            counter = collections.Counter(h["text"] for h in a["headers"] if h["page"] % 2 == parity)
            if not counter:
                continue
            classes = set()
            for t, c in counter.most_common(3):
                classes |= classify_header(t, name, abbrev, title_words, author_words)
            if len(classes) > 1:
                classes.discard("other")
            kinds.append(tuple(sorted(classes)))
            evid.append(ev(a, None, [t for t, c in counter.most_common(2)]))
        v = field(kinds, evid, floor=0.5)
        if v["value"] is not None:
            v["value"] = list(v["value"])
        return v

    def number_position():
        vals, evid = [], []
        for a in articles:
            pos = []
            for n in a["numbers"]:
                if not n.get("width") or not n.get("height") or n.get("y0") is None or n.get("x0") is None or n.get("x1") is None:
                    continue
                if not re.fullmatch(r"[\divxlc]+", n["text"].strip().lower()):
                    continue
                vertical = "top" if n["y0"] < n["height"] * 0.2 else ("bottom" if n["y0"] > n["height"] * 0.8 else "middle")
                center = (n["x0"] + n["x1"]) / 2
                if abs(center - n["width"] / 2) < n["width"] * 0.05:
                    horiz = "center"
                else:
                    horiz = "outer" if (n["page"] % 2 == 1) == (center > n["width"] / 2) else "inner"
                pos.append(f"{vertical}-{horiz}")
            if pos:
                vals.append(mode(pos))
                evid.append(ev(a, None, mode(pos)))
        return field(vals, evid, floor=0.5)
    prof["running_heads"] = {"odd_pages": head_summary(1), "even_pages": head_summary(0), "page_number": number_position()}
    # Title block on the article's first page
    t_size, t_align, t_case, a_case, a_align, t_ev = [], [], [], [], [], []
    for a in articles:
        if a["doc_title"]:
            first = a["doc_title"][0]
            t_size.append(quantize(max((t["size"] or 0) for t in a["doc_title"]), 0.5))
            t_align.append(mode([alignment(t["x0"], t["x1"], t["width"]) for t in a["doc_title"]]))
            t_case.append(text_case(" ".join(t["text"] for t in a["doc_title"])))
            t_ev.append(ev(a, first["page"], first["text"][:60]))
        if a["authors"]:
            a_case.append(text_case(" ".join(t["text"] for t in a["authors"])))
            a_align.append(mode([alignment(t["x0"], t["x1"], t["width"]) for t in a["authors"]]))
    prof["title_block"] = {
        "title_size_pt": field(t_size, t_ev, rounding=1, floor=0.5),
        "title_alignment": field(t_align, t_ev, floor=0.5),
        "title_case": field(t_case, t_ev, floor=0.5),
        "author_case": field(a_case, t_ev, floor=0.5),
        "author_alignment": field(a_align, t_ev, floor=0.5),
        "first_pdf_page_is_article": field([a["first_page"] == 1 for a in articles], [ev(a, a["first_page"], None) for a in articles], floor=0.5),
    }
    prof["table_of_contents"] = field([a["toc_pages"] > 0 for a in articles], [ev(a, None, a["toc_pages"]) for a in articles], floor=0.5)
    # Citation conventions from footnote text
    counts = collections.Counter()
    for a in articles:
        text = "\n".join(a["note_lines"])
        for key, rx in CIT_PATTERNS.items():
            counts[key] += len(rx.findall(text))

    def prefer(a_key, b_key, label_a, label_b):
        a_n, b_n = counts[a_key], counts[b_key]
        if a_n + b_n == 0:
            return {"value": None, "counts": {label_a: 0, label_b: 0}, "confidence": 0.0}
        win = label_a if a_n >= b_n else label_b
        return {"value": win, "counts": {label_a: a_n, label_b: b_n}, "confidence": round(max(a_n, b_n) / (a_n + b_n), 2)}
    prof["citations"] = {
        "ibid_case": prefer("ibid_capitalized", "ibid_lower", "Ibid", "ibid"),
        "ibid_punctuation": prefer("ibid_period", "ibid_comma", "Ibid.", "Ibid,"),
        "supra_form": prefer("supra_note", "supra_comma_note", "supra note N", "supra, note N"),
        "paragraph_pinpoint": prefer("at_para", "at_para_period", "at para N", "at para. N"),
        "page_pinpoint_label": prefer("at_page", "at_p", "at N", "at p N"),
        "range_dash": prefer("range_endash", "range_hyphen", "en dash", "hyphen"),
        "emphasis_note": prefer("emphasis_brackets", "emphasis_parens", "[emphasis added]", "(emphasis added)"),
        "section_abbreviation": prefer("section_s_bare", "section_s_period", "s N", "s. N"),
        "et_al": prefer("et_al_bare", "et_al_period", "et al", "et al."),
        "eg": prefer("eg_bare", "eg_period", "eg", "e.g."),
        "ie": prefer("ie_bare", "ie_period", "ie", "i.e."),
        "see_signal_case": prefer("see_signal_upper", "see_signal_lower", "See", "see"),
        "online_angle_brackets": prefer("online_angle", "online_colon", "online: <url>", "online: url"),
        "quotes": prefer("curly_double", "straight_double", "curly", "straight"),
        "raw_counts": dict(counts),
        "footnote_lines": sum(len(a["note_lines"]) for a in articles),
    }
    prof["status"] = "ok"
    return prof


# ---------------------------------------------------------------- driver ----


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=DB_PATH)
    ap.add_argument("--root", default=OAJD_ROOT)
    ap.add_argument("--out", default=os.path.join("internal", "journal", "profiles"))
    ap.add_argument("--max-articles", type=int, default=10)
    ap.add_argument("--datasets", default="")
    args = ap.parse_args()
    db = sqlite3.connect(f"file:{args.db}?mode=ro", uri=True)
    db.row_factory = sqlite3.Row
    registry = {r["dataset"]: dict(r) for r in db.execute("select * from journal_registry")}
    datasets = [r[0] for r in db.execute(
        "select distinct dataset from articles where document_date_en in (?, ?) order by dataset", YEARS)]
    if args.datasets:
        datasets = [d for d in datasets if d in args.datasets.split(",")]
    os.makedirs(args.out, exist_ok=True)
    index = []
    for ds in datasets:
        rows = db.execute("""
            select a.article_id, a.doc_type, a.name_en, c.source_dir, c.payload_json
            from articles a join article_final_contracts c on c.article_id = a.article_id
            where a.dataset = ? and a.document_date_en in (?, ?)""", (ds,) + YEARS).fetchall()
        cands = []
        for r in rows:
            try:
                payload = json.loads(r["payload_json"] or "{}")
            except Exception:
                payload = {}
            pages = payload.get("page_count") or 0
            doc_type = (r["doc_type"] or "").lower()
            weight = pages
            if any(t in doc_type for t in ("review", "editorial", "note", "comment", "preface", "front", "matter", "full issue", "corrigendum")):
                weight = pages * 0.3
            cands.append((weight, pages, r))
        cands.sort(key=lambda t: -t[0])
        chosen = [c for c in cands if c[1] >= 8][: args.max_articles] or cands[: args.max_articles]
        arts = []
        for _, pages, r in chosen:
            art = analyze_article(ds, r["article_id"], os.path.join(args.root, r["source_dir"]), r["doc_type"])
            if art:
                arts.append(art)
        prof = aggregate(ds, registry.get(ds, {}), arts)
        prof["candidates"] = len(rows)
        path = os.path.join(args.out, f"{ds}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(prof, f, indent=1, ensure_ascii=False)
        page = prof.get("page") or {}
        body = prof.get("body") or {}
        notes = prof.get("footnotes") or {}
        print(f"{ds:26} {len(arts):2}/{len(rows):3} page {page.get('width_in')}x{page.get('height_in')} in; body {body.get('font', {}).get('value')} {body.get('size_pt', {}).get('value')}/{body.get('line_height_pt', {}).get('value')} pt; notes {notes.get('size_pt', {}).get('value')} pt; heads {prof.get('headings', {}).get('numbering_levels', {}).get('value')}")
        index.append({"dataset": ds, "articles": len(arts), "path": os.path.basename(path)})
    with open(os.path.join(args.out, "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, indent=1)


if __name__ == "__main__":
    main()
