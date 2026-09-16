"""Deterministic layout *inputs* and shared-style effect analysis, not rendering.

No line-count proxy, guessed unit, heuristic heading role, or claim of a native
pagination engine. The callers must provide all relevant package story parts to
make a package-wide ownership claim; omitted parts are not certified.
"""
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from typing import Iterable, Mapping
from ooxml import Styles, parse_xml, q, attr


def twips(value: str, unit: str) -> int:
    factors = {'in': Decimal(1440), 'pt': Decimal(20), 'cm': Decimal(1440)/Decimal('2.54')}
    if unit not in factors: raise ValueError('an explicit physical unit is required')
    number = Decimal(value)
    if not number.is_finite(): raise ValueError('non-finite geometry')
    return int((number*factors[unit]).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def affected_styles(styles: Styles, changed: str) -> frozenset[str]:
    if changed not in styles.items: raise ValueError('unknown style')
    affected = {changed}
    # basedOn descendants definitely inherit; linked style partners are included
    # conservatively because native Word's style editing may affect the partner.
    while True:
        before = set(affected)
        for key, node in styles.items.items():
            base = attr(node.child('basedOn'), 'val'); linked = attr(node.child('link'), 'val')
            if base in affected or linked in affected: affected.add(key)
            if key in affected and linked:
                if linked not in styles.items: raise ValueError('unresolved linked style')
                affected.add(linked)
        if affected == before: return frozenset(affected)


@dataclass(frozen=True)
class StyleUse:
    part: str
    element_index: int
    element_kind: str
    style: str


def style_uses(parts: Mapping[str, bytes], styles: Styles) -> tuple[StyleUse, ...]:
    result = []
    for part, data in sorted(parts.items()):
        root = parse_xml(data)
        for i, n in enumerate((root, *root.descendants())):
            if n.tag == q('p'): kind, prop, child = 'paragraph', 'pPr', 'pStyle'
            elif n.tag == q('r'): kind, prop, child = 'character', 'rPr', 'rStyle'
            elif n.tag == q('tbl'): kind, prop, child = 'table', 'tblPr', 'tblStyle'
            else: continue
            pp = n.child(prop)
            key = attr(pp.child(child) if pp else None, 'val', styles.defaults.get(kind, ''))
            if key: styles.chain(key, kind)  # unknown/cyclic references do not disappear
            result.append(StyleUse(part, i, kind, key))
    return tuple(result)


def require_style_ownership(parts: Mapping[str, bytes], styles: Styles, changed: str,
                            owned: frozenset[tuple[str, int]]) -> tuple[StyleUse, ...]:
    affected = affected_styles(styles, changed)
    hits = tuple(u for u in style_uses(parts, styles) if u.style in affected)
    if any((u.part, u.element_index) not in owned for u in hits):
        raise ValueError('shared style affects text outside the declared ownership set')
    return hits


@dataclass(frozen=True)
class Heading:
    id: str
    level: int
    introduction: bool = False


def _roman(n: int) -> str:
    if not 0 < n < 4000: raise ValueError('Roman label outside supported range')
    out = ''
    for value, symbol in ((1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')):
        times, n = divmod(n, value); out += symbol*times
    return out


def _letters(n: int) -> str:
    out = ''
    while n:
        n, digit = divmod(n-1, 26); out = chr(65+digit)+out
    return out


def heading_labels(headings: Iterable[Heading], *, numbered_introduction: bool) -> dict[str, str]:
    """Labels belong to stable supplied heading IDs, never to title-string Find."""
    headings = tuple(headings)
    if len({h.id for h in headings}) != len(headings) or any(not h.id for h in headings):
        raise ValueError('unique nonempty heading identities required')
    counts = [0, 0, 0]; previous = 0; labels = {}
    for i, h in enumerate(headings):
        if not 1 <= h.level <= 3: raise ValueError('only the mapped first three levels are supported')
        if h.introduction and (i != 0 or h.level != 1): raise ValueError('Introduction must be the first level-one heading')
        if h.introduction and not numbered_introduction:
            labels[h.id] = ''; previous = 0; continue
        if h.level > previous + 1: raise ValueError('missing parent heading')
        counts[h.level-1] += 1
        for level in range(h.level, 3): counts[level] = 0
        labels[h.id] = '.'.join((_roman(counts[0]), _letters(counts[1]), str(counts[2]))[:h.level])
        previous = h.level
    return labels
