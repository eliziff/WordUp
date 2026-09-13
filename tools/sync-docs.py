"""Refresh vendored documentation; reruns use locked commits unless --update is given."""
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import sys
import time
import urllib.error
import urllib.request
import zipfile

ROOT = Path(__file__).resolve().parents[1] / "docs" / "upstream"
SOURCES = {
    "vba-docs": ("MicrosoftDocs/VBA-Docs", "main"),
    "open-xml-sdk": ("dotnet/Open-XML-SDK", "v3.3.0"),
    "flaui": ("FlaUI/FlaUI", "v4.0.0"),
    "oletools": ("decalage2/oletools", "v0.60.2"),
    "rubberduck": ("rubberduck-vba/Rubberduck", "fae50adab188126a5e7d2a1cefc3328cc18af482"),
    "ribbonx-editor": ("fernandreu/office-ribbonx-editor", "master"),
    "antlr-go": ("antlr4-go/antlr", "v4.13.1"),
    "pyopenvba": ("WilliamSmithEdward/pyOpenVBA", "main"),
}


def fetch(url):
    request = urllib.request.Request(url, headers={"User-Agent": "WordUp-documentation-sync"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return response.read()
        except urllib.error.URLError:
            if attempt == 2:
                raise
            time.sleep(1)


def selected(source, path):
    p = path.lower()
    name = PurePosixPath(p).name
    if name.startswith(("license", "copying", "notice", "copyright")):
        return True
    if source == "vba-docs":
        return p.endswith((".md", ".yml", ".png", ".jpg", ".svg")) and (
            "/" not in p or p.startswith(("word/", "language/", "library-reference/", "office/", "includes/", "images/"))
            or p.startswith(("api/word.", "api/office.", "api/overview/word")))
    return p.endswith((".md", ".rst")) or (
        p.startswith(("docs/", "doc/")) and p.endswith((".png", ".jpg", ".svg")))


def main():
    ROOT.mkdir(parents=True, exist_ok=True)
    lock_path = ROOT / "sources.json"
    lock = json.loads(lock_path.read_text()) if lock_path.exists() else {}
    if "--check" in sys.argv:
        assert set(lock) == set(SOURCES), "documentation sources missing"
        for name, source in lock.items():
            for p, h in source["files"].items():
                assert hashlib.sha256((ROOT / name / p).read_bytes()).hexdigest() == h, f"changed: {name}/{p}"
        print("All documentation snapshots match their recorded hashes.")
        return
    for name, (repo, ref) in SOURCES.items():
        old = lock.get(name, {})
        if "--update" not in sys.argv and "--refresh" not in sys.argv and old.get("files") and all(
            (ROOT / name / p).is_file() and hashlib.sha256((ROOT / name / p).read_bytes()).hexdigest() == h
            for p, h in old["files"].items()
        ):
            print(f"{name}: verified local snapshot", flush=True)
            continue
        commit = old.get("commit") if "--update" not in sys.argv else None
        if not commit:
            commit = json.loads(fetch(f"https://api.github.com/repos/{repo}/commits/{ref}"))["sha"]
        archive = fetch(f"https://codeload.github.com/{repo}/zip/{commit}")
        files = {}
        with zipfile.ZipFile(io.BytesIO(archive)) as z:
            for member in z.infolist():
                path = PurePosixPath(member.filename)
                relative = PurePosixPath(*path.parts[1:])
                if member.is_dir() or ".." in relative.parts or not selected(name, str(relative)):
                    continue
                data = z.read(member)
                target = ROOT / name / relative
                if not target.resolve().is_relative_to((ROOT / name).resolve()):
                    raise ValueError("unsafe archive path")
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
                files[str(relative)] = hashlib.sha256(data).hexdigest()
        if not any(PurePosixPath(p).name.lower().startswith(("license", "copying")) for p in files):
            raise RuntimeError(f"{name}: no license found; inspect upstream terms")
        for stale in set(old.get("files", {})) - files.keys():
            target = (ROOT / name / stale).resolve()
            if not target.is_relative_to((ROOT / name).resolve()):
                raise ValueError("unsafe stale path")
            target.unlink(missing_ok=True)
        lock[name] = {"repository": f"https://github.com/{repo}", "ref": ref,
                      "commit": commit, "archive_sha256": hashlib.sha256(archive).hexdigest(),
                      "files": files}
        lock_path.write_text(json.dumps(lock, indent=2) + "\n", encoding="utf-8")
        print(f"{name}: {len(files)} files at {commit}", flush=True)


if __name__ == "__main__":
    main()
