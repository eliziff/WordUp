"""Reference binding from complete parsed citations; no fuzzy source matching.

Native note IDs and displayed labels are separate inputs. This model never
inserts a Word field. A pointer-preserving command returns the exact target ID;
a correction command requires a unique parsed source/alias association.
"""
from __future__ import annotations
from dataclasses import dataclass
from hashlib import sha256
from typing import Iterable
from citations import Citation, LegalLocator


@dataclass(frozen=True)
class Note:
    id: str
    label: str
    scope: str
    citations: tuple[Citation, ...]


@dataclass(frozen=True)
class Binding:
    note_id: str
    occurrence: int
    source: tuple[str, ...]
    basis: str | None
    locator: tuple | LegalLocator | None
    qualifications: tuple[str, ...]
    origin_note: str

    @property
    def meaning(self):
        return self.source, self.basis, self.locator, self.qualifications


@dataclass(frozen=True)
class Graph:
    fingerprint: str
    notes: tuple[Note, ...]
    bindings: dict[tuple[str, int], Binding]
    unresolved: dict[tuple[str, int], str]
    aliases: dict[str, tuple[tuple[str, ...], ...]]

    def preserve_pointer(self, scope: str, label: str) -> str:
        targets = [n.id for n in self.notes if n.scope == scope and n.label == label]
        if len(targets) != 1: raise ValueError('display label has no unique native target')
        return targets[0]

    def correct_alias_target(self, note_id: str, alias: str) -> str:
        """Find first preceding *full* citation of one exact declared source.

        Corrects a source association only, not legal relevance. Names/editions
        not connected by a parsed identity remain different. No same-surname rule.
        """
        sources = self.aliases.get(alias, ())
        if len(sources) != 1: raise ValueError('unknown/colliding alias')
        index = next(i for i, n in enumerate(self.notes) if n.id == note_id)
        for n in self.notes[:index]:
            if any(c.source == sources[0] for c in n.citations): return n.id
        raise ValueError('no preceding full-source declaration')


def compile_graph(notes: Iterable[Note]) -> Graph:
    notes = tuple(notes)
    if not notes or any(not n.id or not n.scope or not n.label for n in notes):
        raise ValueError('complete native note identities and numbering scopes required')
    if len({n.id for n in notes}) != len(notes): raise ValueError('duplicate native note ID')
    aliases = {}
    for n in notes:
        for c in n.citations:
            if c.source is not None and c.kind not in {'case','reporter','journal','book','web','legislation'}:
                raise ValueError('source key attached to an unsupported citation kind')
            if c.source is not None and (len(c.source) < 2 or any(not isinstance(s, str) for s in c.source)):
                raise ValueError('malformed parsed source key')
            if c.source is not None and c.declared_alias:
                aliases.setdefault(c.declared_alias, set()).add(c.source)
    aliases = {k: tuple(sorted(v)) for k, v in aliases.items()}
    bindings = {}; unresolved = {}
    g = Graph(sha256(repr(notes).encode()).hexdigest(), notes, bindings, unresolved, aliases)
    index = {n.id: i for i, n in enumerate(notes)}
    for i, n in enumerate(notes):
        for j, c in enumerate(n.citations):
            key = (n.id, j)
            try:
                if c.source is not None:
                    if not c.source[0] or not any(c.source[1:]): raise ValueError('empty source identity')
                    bindings[key] = Binding(n.id, j, c.source, c.basis, c.locator, c.qualifications, n.id)
                    continue
                if c.kind == 'supra':
                    target = g.preserve_pointer(n.scope, c.note_label or '')
                    if index[target] >= i: raise ValueError('supra does not target a preceding note')
                    candidates = [b for (nid, _), b in bindings.items() if nid == target]
                    # An unresolved occurrence in the target makes it incomplete.
                    if len(candidates) != len(notes[index[target]].citations):
                        raise ValueError('target note has unresolved source components')
                    if c.alias:
                        source_set = aliases.get(c.alias, ())
                        if len(source_set) != 1: raise ValueError('unknown/colliding alias')
                        candidates = [b for b in candidates if b.source == source_set[0]]
                    unique = {b.source for b in candidates}
                    if len(unique) != 1: raise ValueError('supra target source is ambiguous')
                    parent = candidates[0]
                    # Supra does not silently inherit the old pinpoint.
                    bindings[key] = Binding(n.id, j, parent.source, c.basis, c.locator,
                                            c.qualifications, parent.origin_note)
                elif c.kind == 'ibid':
                    # Restrict to first citation in current note and exactly one
                    # fully resolved citation in the immediately preceding note.
                    if j != 0 or i == 0 or len(notes[i-1].citations) != 1:
                        raise ValueError('ibid governing citation not uniquely established')
                    parent = bindings.get((notes[i-1].id, 0))
                    if parent is None: raise ValueError('ibid parent unresolved')
                    basis = c.basis if c.locator is not None else parent.basis
                    locator = c.locator if c.locator is not None else parent.locator
                    bindings[key] = Binding(n.id, j, parent.source, basis, locator,
                                            c.qualifications, parent.origin_note)
                else: raise ValueError('unsupported unresolved root')
            except ValueError as e:
                unresolved[key] = str(e)
    return g


def _locator(b: Binding) -> str:
    if b.locator is None: return ''
    if b.basis in ('page', 'paragraph'):
        vals = []
        for item in b.locator:
            if not isinstance(item, tuple) or len(item) != 2 or not all(type(x) is int and x > 0 for x in item):
                raise ValueError('unsupported numeric binding')
            a, z = item
            if z < a: raise ValueError('descending binding')
            vals.append(str(a) if a == z else f'{a}–{z}')
        label = '' if b.basis == 'page' else ('paras ' if len(vals) > 1 or any(a != z for a, z in b.locator) else 'para ')
        return ' at ' + label + ', '.join(vals)
    if b.basis == 'section' and isinstance(b.locator, LegalLocator):
        loc = b.locator
        if len(loc.separators) != len(loc.identifiers) - 1: raise ValueError('incomplete legal locator')
        label = {'section': ('s', 'ss'), 'article': ('art', 'arts'), 'rule': ('rule', 'rules')}[loc.unit][len(loc.identifiers) > 1]
        value = loc.identifiers[0]
        for sep, ident in zip(loc.separators, loc.identifiers[1:]): value += sep + ident
        return ', ' + label + ' ' + value
    raise ValueError('locator needs a lossless renderer')


def repair_ibid_after_reorder(graph: Graph, current: Iterable[Note], note_id: str, *,
                              current_graph_fingerprint: str) -> str | None:
    """Use the old binding, never reinterpret Ibid from its new neighbour.

    A caller provides a validated new native-note ordering/labels, and a fresh
    fingerprint of the old analyzed graph. Added unrelated notes are permitted;
    modifications to any retained citation invalidate this routine. None means
    the existing Ibid still has the exact tested meaning. Text result is a
    proposal, not a native mutation or an automatic acceptance of a note move.
    """
    if current_graph_fingerprint != graph.fingerprint: raise ValueError('stale reference graph')
    current = tuple(current)
    new = compile_graph(current)
    old_by = {n.id: n for n in graph.notes}
    for n in current:
        if n.id in old_by and n.citations != old_by[n.id].citations:
            raise ValueError('a retained citation changed; rebuild associations')
    old = graph.bindings.get((note_id, 0))
    if old is None: raise ValueError('old intended source is unresolved')
    original = old_by[note_id]
    if len(original.citations) != 1 or original.citations[0].kind != 'ibid':
        raise ValueError('not a single bound ibid note')
    index = next(i for i, n in enumerate(current) if n.id == note_id)
    rebound = new.bindings.get((note_id, 0))
    if rebound and rebound.meaning == old.meaning: return None
    origins = [n for n in current[:index] if any(c.source == old.source for c in n.citations)]
    if not origins: raise ValueError('original source has no preceding full citation')
    target = origins[0]
    if not target.label.isascii() or not target.label.isdecimal(): raise ValueError('supra label is not ordinary Arabic numbering')
    if new.preserve_pointer(current[index].scope, target.label) != target.id:
        raise ValueError('cross-scope numbering is ambiguous')
    aliases = sorted(k for k, v in graph.aliases.items() if v == (old.source,))
    if not aliases: raise ValueError('no unique declared short form for old source')
    suffix = ' [' + ', '.join(old.qualifications) + ']' if old.qualifications else ''
    return aliases[0] + ', supra note ' + target.label + _locator(old) + suffix + '.'
