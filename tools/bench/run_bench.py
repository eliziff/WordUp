#!/usr/bin/env python3
"""WordUp agent benchmark runner.

Runs one coding agent (Codex CLI, gpt-5.6-luna) against template-editing tasks
under two tool profiles, then grades the produced workspace with hidden native
Word suites that the agent never sees. Nothing here estimates dollar cost: the
Codex path is subscription-backed, so tokens and wall time are the budget.

    python tools/bench/run_bench.py --arms core,components --runs 3
    python tools/bench/run_bench.py --arms none            # grade fixtures untouched (control)
    python tools/bench/run_bench.py --task alr-* --exe bin/wordup.exe --out tools/bench/results

Each run lives under results/<stamp>/<task>/<arm>/run-N/ with the disposable
workspace copy, the agent's AGENTS.md, the Codex JSONL event stream, the final
message, the grade, and summary.json. results/<stamp>/summary.md aggregates.
"""
from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import glob
import json
import os
import pathlib
import shutil
import statistics
import subprocess
import sys
import time

HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parents[1]

ARMS = {
    "none": {
        "profile": "full",
        "note": "No agent runs; the fixture is graded untouched as a control.",
    },
    "core": {
        "profile": "core",
        "note": (
            "This tool profile exposes only WordUp's core surface. The bundled VBA "
            "component library (component.*), journal generation (journal.*) and "
            "structure detection (structure.*) are DISABLED: write any helper VBA you "
            "need yourself inside the workspace."
        ),
    },
    "components": {
        "profile": "full",
        "note": (
            "Reusable, natively tested VBA components are available: "
            "`tool\\wordup.cmd -w workspace call component.list {}` lists them, "
            "`call component.get {\"component\":\"ID\"}` shows a component's source, and "
            "`call component.add {\"component\":\"ID\"}` installs it into workspace\\vba. "
            "Use them where they fit; you may also write your own VBA."
        ),
    },
}

AGENTS_TEMPLATE = """# Benchmark task

You are working in a disposable copy of a WordUp template workspace at `workspace\\`.
Complete the task below end to end, then stop. Do not ask questions; nobody will answer.
Do not read or modify anything outside this directory.

## Tool

Always run WordUp through the wrapper `tool\\wordup.cmd` (it pins the tool profile and a local TEMP).
Read `tool\\docs\\API.md` (operations, acceptance suites), `tool\\docs\\LIMITS.md`, and `workspace\\AGENTS.md`.

Useful commands (run from this directory):

    tool\\wordup.cmd help
    tool\\wordup.cmd doctor --word
    tool\\wordup.cmd -w workspace check
    tool\\wordup.cmd -w workspace build
    tool\\wordup.cmd -w workspace --execute session start
    tool\\wordup.cmd -w workspace rpc native.call @operation.json
    tool\\wordup.cmd -w workspace --execute test workspace\\dist\\<artifact>.dotm workspace\\tests\\<suite>.json
    tool\\wordup.cmd -w workspace session stop

Microsoft Word is installed; native execution through the tool is expected and authorized for this task.
Write your own acceptance suite under `workspace\\tests\\` and run it in real Word before you finish.
`build` is not a compile: use `test` (with a compile step) or `native.call` to prove the VBA runs.

{arm_note}

## Task

{prompt}

## Finishing

Make sure `tool\\wordup.cmd -w workspace build` succeeds at the end. Your final message must be JSON
matching the output schema: summary, files_changed, components_used, native_tests_run, confidence.
"""

WRAPPER = """@echo off
set WORDUP_TOOL_PROFILE={profile}
set TEMP={tmp}
set TMP={tmp}
"%~dp0wordup.exe" %*
"""

DOCS = ["README.md", "AGENTS.md", "docs/API.md", "docs/LIMITS.md", "docs/DIAGNOSTICS.md",
        "docs/WORD-OPERATIONS.md", "docs/EXPECTED-RESULTS.md"]

COPY_IGNORE = ("dist", "reports", "build", ".wordup", "acceptance-assets", "__pycache__", "*.pyc")


def log(msg: str) -> None:
    print(time.strftime("%H:%M:%S"), msg, flush=True)


def run(cmd: list[str], cwd: pathlib.Path | None = None, timeout: int | None = None,
        stdin_text: str | None = None, env: dict | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=str(cwd) if cwd else None, input=stdin_text, capture_output=True,
                          text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env)


def parse_json_output(text: str) -> dict:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        if start >= 0:
            try:
                return json.loads(text[start:])
            except json.JSONDecodeError:
                pass
    return {"raw": text[-4000:]}


def word_process_count() -> int:
    if os.name != "nt":
        return -1
    p = run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"])
    return sum(1 for line in p.stdout.splitlines() if line.startswith('"WINWORD.EXE"'))


def find_codex() -> str:
    for candidate in (shutil.which("codex"), shutil.which("codex.cmd")):
        if candidate:
            return candidate
    appdata = os.environ.get("APPDATA", "")
    for name in ("codex.cmd", "codex.exe"):
        path = os.path.join(appdata, "npm", name)
        if os.path.exists(path):
            return path
    raise SystemExit("codex CLI not found on PATH")


def load_tasks(pattern: str) -> list[dict]:
    tasks = []
    for path in sorted((HERE / "tasks").glob("*.json")):
        task = json.loads(path.read_text(encoding="utf-8"))
        task["_path"] = path
        if fnmatch.fnmatch(task["id"], pattern):
            tasks.append(task)
    if not tasks:
        raise SystemExit(f"no tasks match {pattern!r}")
    return tasks


def prepare_fixture(exe: pathlib.Path, task: dict, workspace: pathlib.Path) -> None:
    fixture = task["fixture"]
    kind = fixture.get("type", "workspace")
    if kind == "workspace":
        source = pathlib.Path(fixture["path"])
        if not source.is_absolute():
            source = REPO / source
        if not source.is_dir():
            raise SystemExit(f"fixture workspace missing: {source}")
        shutil.copytree(source, workspace, ignore=shutil.ignore_patterns(*COPY_IGNORE))
    elif kind == "example":
        p = run([str(exe), "example", str(workspace)])
        if p.returncode != 0:
            raise SystemExit(f"example fixture failed: {p.stdout[-2000:]} {p.stderr[-2000:]}")
        for stale in ("dist", "reports"):
            shutil.rmtree(workspace / stale, ignore_errors=True)
    else:
        raise SystemExit(f"unknown fixture type {kind}")


def stage_tool(exe: pathlib.Path, run_root: pathlib.Path, profile: str) -> pathlib.Path:
    tool = run_root / "tool"
    tool.mkdir(parents=True, exist_ok=True)
    shutil.copy2(exe, tool / "wordup.exe")
    docs = tool / "docs"
    docs.mkdir(exist_ok=True)
    for rel in DOCS:
        src = REPO / rel
        if src.exists():
            shutil.copy2(src, docs / src.name)
    tmp = run_root / "tmp"
    tmp.mkdir(exist_ok=True)
    (tool / "wordup.cmd").write_text(WRAPPER.format(profile=profile, tmp=str(tmp)), encoding="ascii")
    return tool


def summarize_events(events_path: pathlib.Path) -> dict:
    usage = {"input_tokens": 0, "cached_input_tokens": 0, "output_tokens": 0, "reasoning_output_tokens": 0}
    commands, messages, turns = 0, 0, 0
    last_usage = None
    if not events_path.exists():
        return {"usage": usage, "commands": 0, "messages": 0, "turns": 0}
    for line in events_path.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        kind = event.get("type", "")
        item = event.get("item") or {}
        if kind == "item.started" and item.get("type") == "command_execution":
            commands += 1
        if kind == "item.completed" and item.get("type") == "agent_message":
            messages += 1
        if kind == "turn.completed":
            turns += 1
            last_usage = event.get("usage")
        if isinstance(event.get("usage"), dict):
            last_usage = event["usage"]
    if isinstance(last_usage, dict):
        for key in usage:
            value = last_usage.get(key)
            if isinstance(value, (int, float)):
                usage[key] = int(value)
    return {"usage": usage, "commands": commands, "messages": messages, "turns": turns}


def run_agent(codex: str, run_root: pathlib.Path, prompt: str, args, profile: str) -> dict:
    events = run_root / "events.jsonl"
    last = run_root / "last.md"
    cmd = [codex, "exec", "--color", "never", "--json", "-o", str(last), "-C", str(run_root),
           "--skip-git-repo-check", "--ephemeral", "--ignore-user-config", "--ignore-rules",
           "-m", args.model, "-c", f'model_reasoning_effort="{args.effort}"',
           "--output-schema", str(HERE / "schemas" / "result.json")]
    if args.sandbox == "danger-full-access":
        # Word must write its own profile directories, so the OS sandbox is off;
        # the run directory is disposable and fixtures are copies.
        cmd.append("--dangerously-bypass-approvals-and-sandbox")
    else:
        cmd += ["--sandbox", args.sandbox, "-c", 'approval_policy="never"', "--add-dir", str(run_root / "tmp")]
    cmd.append("-")
    env = dict(os.environ)
    env["WORDUP_TOOL_PROFILE"] = profile
    started = time.time()
    timed_out = False
    try:
        with open(events, "w", encoding="utf-8") as out, open(run_root / "agent-stderr.txt", "w", encoding="utf-8") as err:
            proc = subprocess.Popen(cmd, cwd=str(run_root), stdin=subprocess.PIPE, stdout=out, stderr=err,
                                    text=True, encoding="utf-8", env=env)
            proc.stdin.write(prompt)
            proc.stdin.close()
            try:
                proc.wait(timeout=args.timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                proc.kill()
                proc.wait()
    finally:
        wall = time.time() - started
    result = {"wall_s": round(wall, 1), "timed_out": timed_out, "exit_code": None if timed_out else proc.returncode,
              "command": cmd}
    result.update(summarize_events(events))
    if last.exists():
        text = last.read_text(encoding="utf-8", errors="replace")
        result["final_message"] = text[-4000:]
        try:
            result["final_json"] = json.loads(text)
        except json.JSONDecodeError:
            result["final_json"] = None
    return result


def grade(exe: pathlib.Path, run_root: pathlib.Path, task: dict) -> dict:
    workspace = run_root / "workspace"
    spec = task["grade"]
    out = {"checks": [], "passed": False}
    output = workspace / spec.get("build_output", "dist/bench.dotm")
    build = run([str(exe), "-w", str(workspace), "build", str(output)], timeout=600)
    build_json = parse_json_output(build.stdout)
    build_ok = build.returncode == 0 and "error" not in build_json
    out["build"] = {"ok": build_ok, "result": build_json.get("result"), "error": build_json.get("error"),
                    "duration_ms": build_json.get("duration_ms")}
    out["checks"].append({"name": "build", "passed": build_ok})
    artifact = None
    if build_ok:
        artifact = (build_json.get("result") or {}).get("artifact") or str(output)
    for needle in spec.get("check_forbid", []):
        check = run([str(exe), "-w", str(workspace), "check"], timeout=600)
        hit = needle in check.stdout
        out["checks"].append({"name": f"check has no diagnostic mentioning {needle}", "passed": not hit})
    for item in spec.get("file_contains", []):
        matches = glob.glob(str(workspace / item["path"]), recursive=True)
        found = any(item["text"] in pathlib.Path(m).read_text(encoding="utf-8", errors="replace") for m in matches)
        out["checks"].append({"name": f"{item['path']} contains {item['text']!r}", "passed": found})
    suite = spec.get("suite")
    if suite and artifact:
        suite_path = HERE / suite
        before = word_process_count()
        test = run([str(exe), "-w", str(workspace), "--execute", "test", artifact, str(suite_path)], timeout=1800)
        report = parse_json_output(test.stdout)
        result = report.get("result") or {}
        status = result.get("status")
        out["test"] = {"status": status, "assertions": result.get("assertions"), "vba_compiled": result.get("vba_compiled"),
                       "word_executed": result.get("word_executed"), "duration_ms": result.get("duration_ms"),
                       "error": report.get("error") or result.get("error"), "saved_report": result.get("saved_report")}
        (run_root / "grade-test-report.json").write_text(json.dumps(report, indent=1)[:2_000_000], encoding="utf-8")
        out["checks"].append({"name": "hidden native suite", "passed": status == "passed"})
        out["word_processes_before"], out["word_processes_after"] = before, word_process_count()
    elif suite:
        out["checks"].append({"name": "hidden native suite", "passed": False, "reason": "build failed"})
    out["passed"] = all(c["passed"] for c in out["checks"])
    return out


def median(values):
    values = [v for v in values if isinstance(v, (int, float))]
    return round(statistics.median(values), 1) if values else None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--exe", default=str(REPO / "bin" / "wordup.exe"))
    ap.add_argument("--task", default="*", help="glob over task ids")
    ap.add_argument("--arms", default="core,components")
    ap.add_argument("--runs", type=int, default=1)
    ap.add_argument("--out", default=str(HERE / "results"))
    ap.add_argument("--label", default="")
    ap.add_argument("--model", default="gpt-5.6-luna")
    ap.add_argument("--effort", default="xhigh")
    ap.add_argument("--timeout", type=int, default=1800, help="seconds per agent run")
    ap.add_argument("--sandbox", default="danger-full-access", choices=["workspace-write", "danger-full-access", "read-only"],
                    help="Codex sandbox; Word needs its profile directories, so full access is the default")
    ap.add_argument("--keep-workspaces", action="store_true", help="keep the disposable workspace copies (default: delete after grading)")
    args = ap.parse_args()

    exe = pathlib.Path(args.exe).resolve()
    if not exe.exists():
        raise SystemExit(f"wordup executable not found: {exe}")
    arms = [a.strip() for a in args.arms.split(",") if a.strip()]
    for arm in arms:
        if arm not in ARMS:
            raise SystemExit(f"unknown arm {arm}; choose from {', '.join(ARMS)}")
    codex = find_codex() if any(a != "none" for a in arms) else None
    tasks = load_tasks(args.task)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S") + (f"-{args.label}" if args.label else "")
    root = pathlib.Path(args.out) / stamp
    root.mkdir(parents=True)
    log(f"benchmark {stamp}: {len(tasks)} tasks x {arms} x {args.runs} runs; exe {exe}")
    rows = []
    for task in tasks:
        for arm in arms:
            profile = ARMS[arm]["profile"]
            for n in range(1, args.runs + 1):
                run_root = root / task["id"] / arm / f"run-{n}"
                run_root.mkdir(parents=True)
                workspace = run_root / "workspace"
                prepare_fixture(exe, task, workspace)
                tool = stage_tool(exe, run_root, profile)
                agents_md = AGENTS_TEMPLATE.format(arm_note=ARMS[arm]["note"], prompt=task["prompt"])
                (run_root / "AGENTS.md").write_text(agents_md, encoding="utf-8")
                summary = {"task": task["id"], "arm": arm, "run": n, "profile": profile, "exe": str(exe),
                           "exe_sha256": sha256(exe), "started": dt.datetime.now().isoformat(timespec="seconds")}
                if arm != "none":
                    log(f"{task['id']} / {arm} / run-{n}: agent")
                    summary["agent"] = run_agent(codex, run_root, agents_md, args, profile)
                    # Leave no owned Word behind from the agent's own sessions.
                    run([str(tool / "wordup.cmd"), "-w", str(workspace), "session", "stop"], timeout=60)
                log(f"{task['id']} / {arm} / run-{n}: grade")
                summary["grade"] = grade(exe, run_root, task)
                (run_root / "summary.json").write_text(json.dumps(summary, indent=1), encoding="utf-8")
                rows.append(summary)
                log(f"{task['id']} / {arm} / run-{n}: {'PASS' if summary['grade']['passed'] else 'FAIL'}")
                if not args.keep_workspaces:
                    shutil.rmtree(workspace, ignore_errors=True)
                    shutil.rmtree(run_root / "tmp", ignore_errors=True)
    write_summary(root, rows, args)
    log(f"done: {root / 'summary.md'}")
    return 0


def sha256(path: pathlib.Path) -> str:
    import hashlib
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def write_summary(root: pathlib.Path, rows: list[dict], args) -> None:
    groups: dict[tuple, list[dict]] = {}
    for r in rows:
        groups.setdefault((r["task"], r["arm"]), []).append(r)
    lines = [f"# Benchmark {root.name}", "", f"model {args.model} effort {args.effort}; exe {args.exe}; timeout {args.timeout}s per run", "",
             "| task | arm | runs | passed | median wall s | median commands | median input tok | median output tok | median suite ms |",
             "|---|---|---:|---:|---:|---:|---:|---:|---:|"]
    for (task, arm), items in sorted(groups.items()):
        agents = [i.get("agent") or {} for i in items]
        lines.append("| {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
            task, arm, len(items), sum(1 for i in items if i["grade"]["passed"]),
            median([a.get("wall_s") for a in agents]), median([a.get("commands") for a in agents]),
            median([(a.get("usage") or {}).get("input_tokens") for a in agents]),
            median([(a.get("usage") or {}).get("output_tokens") for a in agents]),
            median([(i["grade"].get("test") or {}).get("duration_ms") for i in items])))
    lines += ["", "## Runs", ""]
    for r in rows:
        failed = [c["name"] for c in r["grade"]["checks"] if not c["passed"]]
        agent = r.get("agent") or {}
        lines.append(f"- {r['task']} / {r['arm']} / run-{r['run']}: {'PASS' if r['grade']['passed'] else 'FAIL ' + ', '.join(failed)}"
                     f"; wall {agent.get('wall_s')} s; commands {agent.get('commands')}; timed_out {agent.get('timed_out')}"
                     f"; components {((agent.get('final_json') or {}).get('components_used'))}")
    (root / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    (root / "rows.json").write_text(json.dumps(rows, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
