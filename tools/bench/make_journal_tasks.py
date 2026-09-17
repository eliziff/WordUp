"""Generate journal-generation benchmark tasks from bundled inferred profiles.

For each journal id a fixture workspace is created with `wordup new` and the
journal's profile (as journal.profile returns it, layout and conventions
included) is written to journal/profile.json. The task asks the agent for a
Public Sub WU_Typeset that applies the profile; the hidden suite asserts the
concrete values the profile carries, so the untouched fixture fails and a
template built from the profile passes.

Usage:
    python tools/bench/make_journal_tasks.py --exe build/wordup.exe --journals UBC-L-REV,OSGOODE-HALL-LJ,MCGILL-LJ-HEALTH
"""
from __future__ import annotations

import argparse
import json
import pathlib
import shutil
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parent.parent


def run(args):
    p = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=REPO)
    out = (p.stdout or "").lstrip("﻿")
    try:
        return json.loads(out)
    except Exception:
        raise SystemExit(f"{' '.join(str(a) for a in args)} failed: {out[-800:]} {p.stderr[-800:]}")


def safe_identifier(value: str) -> str:
    out = "".join(ch if ch.isalnum() else "_" for ch in value)
    if not out or out[0].isdigit():
        out = "J_" + out
    return out


def build_suite(profile: dict) -> dict:
    layout = profile.get("layout") or {}
    conventions = profile.get("conventions") or {}
    checks = []
    if layout.get("page_width_pt"):
        checks.append(f'If Abs(d.Sections(1).PageSetup.PageWidth - {layout["page_width_pt"]:.2f}) > 0.5 Or Abs(d.Sections(1).PageSetup.PageHeight - {layout["page_height_pt"]:.2f}) > 0.5 Then Err.Raise 5, , "page size " & d.Sections(1).PageSetup.PageWidth & "x" & d.Sections(1).PageSetup.PageHeight')
    if layout.get("top_in"):
        checks.append(f'If Abs(d.Sections(1).PageSetup.TopMargin - InchesToPoints({layout["top_in"]:.2f})) > 1 Then Err.Raise 5, , "top margin " & d.Sections(1).PageSetup.TopMargin')
    if layout.get("mirror_margins"):
        checks.append('If d.Sections(1).PageSetup.MirrorMargins <> True Then Err.Raise 5, , "mirror margins"')
    elif layout.get("left_in"):
        checks.append(f'If Abs(d.Sections(1).PageSetup.LeftMargin - InchesToPoints({layout["left_in"]:.2f})) > 1 Then Err.Raise 5, , "left margin " & d.Sections(1).PageSetup.LeftMargin')
    body_font = profile.get("body_font")
    body_size = profile.get("body_size_pt")
    if body_font and body_size:
        checks.append(f'If StrComp(d.Paragraphs(4).Range.Font.Name, {json.dumps(body_font)}, vbTextCompare) <> 0 Or Abs(d.Paragraphs(4).Range.Font.Size - {body_size:.2f}) > 0.05 Then Err.Raise 5, , "body font " & d.Paragraphs(4).Range.Font.Name & " " & d.Paragraphs(4).Range.Font.Size')
    note_size = profile.get("note_size_pt")
    if note_size:
        checks.append(f'If Abs(d.Footnotes(1).Range.Font.Size - {note_size:.2f}) > 0.05 Then Err.Raise 5, , "note size " & d.Footnotes(1).Range.Font.Size')
    if layout.get("heading_case") == "upper":
        checks.append('If d.Paragraphs(3).Range.Font.AllCaps <> True And d.Paragraphs(3).Range.Text <> UCase$(d.Paragraphs(3).Range.Text) Then Err.Raise 5, , "heading is not upper case"')
    if layout.get("heading_alignment") == "center":
        checks.append('If d.Paragraphs(3).Alignment <> wdAlignParagraphCenter Then Err.Raise 5, , "heading is not centred"')
    if layout.get("odd_running_head") or layout.get("even_running_head") or layout.get("page_number"):
        checks.append('If Len(Trim$(Replace(Replace(d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text, vbCr, ""), vbTab, ""))) = 0 And d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Fields.Count = 0 And d.Sections(1).Footers(wdHeaderFooterPrimary).Range.Fields.Count = 0 Then Err.Raise 5, , "no running head or page number"')
    if layout.get("note_number_style") in ("number_gap", "number_space"):
        checks.append('If Left$(d.Footnotes(1).Range.Text, 1) <> vbTab Then Err.Raise 5, , "footnote has no tab after its number"')
    if conventions.get("quotes") == "curly":
        checks.append('If InStr(1, d.Paragraphs(4).Range.Text, """", vbBinaryCompare) > 0 Then Err.Raise 5, , "straight quotes survived"')
    checks.append('If StrComp(d.Content.Text, before, vbTextCompare) <> 0 Then Err.Raise 5, , "body text changed: " & d.Content.Text')
    # The seed and the checks are scratch evaluations; the macro itself runs
    # through the harness's runner, which resolves a bare macro name in the
    # opened template where a scratch project's Application.Run cannot.
    seed = "\n".join([
        "Dim d As Document, rng As Range",
        "Set d = ActiveDocument",
        'd.Content.Text = "A Synthetic Article Title" & vbCr & "By Jane Doe" & vbCr & "Introduction" & vbCr & "Body text with ""straight quotes"" in it." & vbCr & "Second paragraph." & vbCr',
        "d.Content.Style = wdStyleNormal",
        "d.Paragraphs(3).OutlineLevel = wdOutlineLevel1",
        "Set rng = d.Paragraphs(4).Range.Duplicate: rng.Collapse wdCollapseEnd: rng.Move wdCharacter, -1",
        'd.Footnotes.Add rng, , "See Smith, ibid at para 12."',
        'd.BuiltInDocumentProperties("Title").Value = "A Synthetic Article Title"',
        'd.Variables("BenchWidthBefore").Value = CStr(d.Sections(1).PageSetup.PageWidth)',
        'd.Variables("BenchTextBefore").Value = Replace(Replace(d.Content.Text, """", ChrW(8220)), "straight quotes" & ChrW(8220), "straight quotes" & ChrW(8221))',
        "Evaluate = d.Footnotes.Count",
    ])
    verify = "\n".join([
        "Dim d As Document, before As String, widthBefore As Single",
        "Set d = ActiveDocument",
        'widthBefore = CSng(d.Variables("BenchWidthBefore").Value)',
        'before = d.Variables("BenchTextBefore").Value',
        *checks,
        "d.Undo",
        'If Abs(d.Sections(1).PageSetup.PageWidth - widthBefore) > 0.5 And ' + ("True" if layout.get("page_width_pt") else "False") + ' Then Err.Raise 5, , "one undo did not restore the page size"',
        "d.Close SaveChanges:=wdDoNotSaveChanges",
        'Evaluate = "PASS"',
    ])
    return {"schema": 1, "name": f"Bench: {profile['id']} typeset from profile", "require_compile": True, "steps": [
        {"name": "Open template", "operation": {"op": "open", "file": "$artifact", "as": "template"}},
        {"name": "Compile template", "operation": {"op": "compile", "target": "template", "member": "$project"}, "assert": [{"path": "/vba_compiled", "kind": "equals", "expected": True}]},
        {"name": "Create document from the template", "operation": {"op": "new", "file": "$artifact", "as": "document"}},
        {"name": "Seed a synthetic manuscript", "operation": {"op": "eval", "value": seed}, "assert": [{"path": "/result", "kind": "equals", "expected": 1}]},
        {"name": "Run WU_Typeset", "operation": {"op": "run", "macro": "WU_Typeset", "timeout_ms": 120000}},
        {"name": "Verify the profile was applied", "operation": {"op": "eval", "value": verify, "timeout_ms": 120000}, "assert": [{"path": "/result", "kind": "equals", "expected": "PASS"}]},
    ]}


def prompt_for(profile: dict) -> str:
    layout = profile.get("layout") or {}
    conventions = profile.get("conventions") or {}
    facts = []
    if layout.get("page_width_pt"):
        facts.append(f"page {layout['page_width_pt']:.0f} x {layout['page_height_pt']:.0f} pt")
    if layout.get("mirror_margins"):
        facts.append(f"mirrored margins inside {layout.get('inside_in')} in, outside {layout.get('outside_in')} in")
    elif layout.get("left_in"):
        facts.append(f"margins left {layout.get('left_in')} in, right {layout.get('right_in')} in")
    if layout.get("top_in"):
        facts.append(f"top {layout['top_in']} in")
    facts.append(f"body {profile.get('body_font')} {profile.get('body_size_pt')} pt, notes {profile.get('note_font')} {profile.get('note_size_pt')} pt")
    if layout.get("heading_case"):
        facts.append(f"top-level headings {layout['heading_case']} case, {layout.get('heading_alignment') or 'left'} aligned")
    if layout.get("odd_running_head") or layout.get("even_running_head"):
        facts.append(f"running heads: odd pages {layout.get('odd_running_head')}, even pages {layout.get('even_running_head')}, page number {layout.get('page_number') or 'not printed in the header'}")
    if layout.get("note_number_style") in ("number_gap", "number_space"):
        facts.append("a tab between the footnote number and its text")
    if conventions.get("quotes") == "curly":
        facts.append("curly quotes")
    return (
        f"This workspace is a blank Word template project for {profile['name']}. journal/profile.json describes how the journal typesets its published "
        f"articles, measured on its 2025-2026 output: {'; '.join(facts)}. Every field carries the evidence; fields listed under evidence_gaps have none and must be left alone. "
        "Write the VBA so that a Public Sub named WU_Typeset (no arguments) applies the profile to the active document in one Application.UndoRecord custom record: "
        "page size and margins on every section (write only values that differ), body and footnote fonts and sizes (through styles, not direct formatting), the top-level "
        "heading case and alignment on outline-level-1 paragraphs, the running heads and page-number field the profile describes (leave a header an editor already filled), "
        "a tab after each footnote number when the profile says so, and curly quotes when it says so. Body text must otherwise stay identical and one Undo must revert everything. "
        "WU_Typeset must run without any user interaction (no message boxes or forms) because it is graded by automation. "
        "Build the template with `tool\\wordup.cmd -w workspace build` and prove WU_Typeset in real Word on a disposable document before you finish."
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=str(REPO / "build" / "wordup.exe"))
    ap.add_argument("--journals", default="UBC-L-REV,OSGOODE-HALL-LJ,MCGILL-LJ-HEALTH")
    args = ap.parse_args()
    exe = pathlib.Path(args.exe)
    if not exe.exists():
        raise SystemExit(f"wordup executable not found: {exe}")
    for journal in [j.strip() for j in args.journals.split(",") if j.strip()]:
        profile = run([str(exe), "call", "journal.profile", json.dumps({"journal": journal})]).get("result")
        if not profile:
            raise SystemExit(f"no profile for {journal}")
        slug = "journal-" + journal.lower()
        fixture = HERE / "fixtures" / slug
        if fixture.exists():
            shutil.rmtree(fixture)
        run([str(exe), "new", "WU_" + safe_identifier(journal), str(fixture)])
        (fixture / "journal").mkdir(exist_ok=True)
        (fixture / "journal" / "profile.json").write_text(json.dumps(profile, indent=1, ensure_ascii=False), encoding="utf-8")
        suite_path = HERE / "suites" / f"{slug}.json"
        suite_path.write_text(json.dumps(build_suite(profile), indent=1, ensure_ascii=False), encoding="utf-8")
        task = {
            "id": slug,
            "title": f"Typeset macro for {profile['name']} from its measured profile",
            "tags": ["journal", "generation", "profile"],
            "fixture": {"type": "workspace", "path": str(fixture.relative_to(REPO)).replace("\\", "/")},
            "prompt": prompt_for(profile),
            "grade": {"build_output": "dist/bench.dotm", "suite": f"suites/{slug}.json"},
            "timeout_s": 1800,
        }
        (HERE / "tasks" / f"{slug}.json").write_text(json.dumps(task, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"{slug}: fixture {fixture.relative_to(REPO)}, suite suites/{slug}.json, {len(build_suite(profile)['steps'][5]['operation']['value'].splitlines())} verify lines")


if __name__ == "__main__":
    sys.exit(main())
