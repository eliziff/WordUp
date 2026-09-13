"""Development-only writer checks. Inputs are never overwritten; evidence stays local."""
import argparse
import copy
import hashlib
import io
import json
import os
from pathlib import Path
import statistics
import subprocess
import sys
import time
import zipfile

PIN = "7aeb36caadf49177a643c490f348be44c625d0e2"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def parts(path):
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        if len(names) != len(set(names)):
            raise ValueError("Duplicate package entries")
        return {name: hashlib.sha256(package.read(name)).hexdigest() for name in names}


def streams(path):
    # Independent binary reader, not the writer's own round-trip interpretation.
    import olefile
    with zipfile.ZipFile(path) as package:
        data = package.read("word/vbaProject.bin")
    with olefile.OleFileIO(io.BytesIO(data)) as compound:
        return {"/".join(name): hashlib.sha256(compound.openstream(name).read()).hexdigest()
                for name in compound.listdir(streams=True, storages=False)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--upstream", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--repetitions", type=int, default=7)
    parser.add_argument("--go-report", type=Path)
    parser.add_argument("inputs", type=Path, nargs="+")
    args = parser.parse_args()
    if not 2 <= args.repetitions <= 100:
        parser.error("repetitions must be 2..100")
    revision = subprocess.check_output(["git", "-C", str(args.upstream), "rev-parse", "HEAD"], text=True).strip()
    if revision != PIN or subprocess.check_output(["git", "-C", str(args.upstream), "status", "--porcelain", "--untracked-files=no"], text=True).strip():
        raise ValueError("Expected unmodified pinned pyOpenVBA source")
    sys.path.insert(0, str(args.upstream / "src"))
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.GetCurrentProcess.restype = wintypes.HANDLE
        kernel.SetPriorityClass.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        if not kernel.SetPriorityClass(kernel.GetCurrentProcess(), 0x4000):
            raise OSError("Cannot set BelowNormal priority")
    started = time.perf_counter()
    from pyopenvba import WordFile
    import_ms = (time.perf_counter() - started) * 1000
    args.output.mkdir(parents=True, exist_ok=False)
    results = []
    for index, source in enumerate(args.inputs):
        original_hash = digest(source)
        original_parts = parts(source)
        original_streams = streams(source)
        with WordFile(source) as document:
            original_modules = document.vba_modules()
        for operation in ("unchanged", "module-edit", "module-lifecycle", "forms-read", "control-caption", "control-font"):
            samples = []
            row = {"input": str(source.resolve()), "source_sha256": original_hash, "operation": operation}
            try:
                for iteration in range(args.repetitions):
                    output = args.output / f"{index}-{operation}-{iteration}.dotm"
                    start = time.perf_counter()
                    with WordFile(source) as document:
                        modules = document.vba_modules()
                        target = sorted(modules)[0]
                        expected = dict(original_modules)
                        if operation == "module-edit":
                            expected[target] = modules[target].rstrip("\r\n") + "\r\n' WordUp writer comparison\r\n"
                            document.set_module(target, expected[target])
                        if operation == "module-lifecycle":
                            project = document.vba_project()
                            added = 'Attribute VB_Name = "WriterAdded"\r\nPublic Function WriterAnswer() As Long\r\nWriterAnswer = 42\r\nEnd Function\r\n'
                            project.add_module("WriterAdded", added)
                            project.add_module("WriterRemoved", 'Attribute VB_Name = "WriterRemoved"\r\nPublic Sub Temporary()\r\nEnd Sub\r\n')
                            document.save(args.output / f"{index}-lifecycle-intermediate-{iteration}.dotm")
                            project.rename_module("WriterAdded", "WriterRenamed")
                            project.delete_module("WriterRemoved")
                            expected["WriterRenamed"] = added.replace('VB_Name = "WriterAdded"', 'VB_Name = "WriterRenamed"')
                        if operation == "forms-read":
                            row["forms"] = len(document.forms())
                        expected_controls = None
                        if operation in ("control-caption", "control-font"):
                            controls = {(form.name, control.name): control for form in document.forms() for control in form.walk()}
                            expected_controls = {key: (control.kind, control.id, control.tab_index, copy.deepcopy(control.properties()))
                                                 for key, control in controls.items()}
                            eligible = sorted(key for key, control in controls.items() if control.kind in ("MSForms.Label", "MSForms.CommandButton"))
                            if not eligible:
                                raise ValueError("Precondition unmet: no supported caption control")
                            selected = eligible[0]
                            if operation == "control-caption":
                                controls[selected].set_property("Caption", "Writer comparison")
                                expected_controls[selected][3]["Caption"] = "Writer comparison"
                            else:
                                font = controls[selected].record.text_props
                                if font is None:
                                    raise ValueError("Precondition unmet: control has no TextProps record")
                                font.set_string("FontName", "Segoe UI")
                                font.set_value("FontHeight", 240)
                                expected_controls[selected][3].update({"Font.FontName":"Segoe UI", "Font.FontHeight":240})
                            row["control"] = list(selected)
                        document.save(output)
                    samples.append((time.perf_counter() - start) * 1000)
                    with WordFile(output) as document:
                        normalize = lambda values: {key: value.replace("\r\n", "\n").rstrip("\n") for key, value in values.items()}
                        if normalize(document.vba_modules()) != normalize(expected):
                            raise AssertionError("Module source preservation failed")
                        if expected_controls is not None:
                            actual_controls = {(form.name, control.name): (control.kind, control.id, control.tab_index, control.properties())
                                               for form in document.forms() for control in form.walk()}
                            if actual_controls != expected_controls:
                                raise AssertionError("Control property or inventory preservation failed")
                    output_parts = parts(output)
                    changed = sorted(name for name in set(original_parts) | set(output_parts) if original_parts.get(name) != output_parts.get(name))
                    output_streams = streams(output)
                    stream_changes = sorted(name for name in set(original_streams) | set(output_streams)
                                            if original_streams.get(name) != output_streams.get(name))
                    row.update(output=str(output.resolve()), output_sha256=digest(output), changed_parts=changed,
                               changed_streams=stream_changes)
                    mutation = operation in ("module-edit", "module-lifecycle", "control-caption", "control-font")
                    allowed = {"word/vbaProject.bin"} if mutation else set()
                    if set(changed) - allowed:
                        raise AssertionError("Unexpected changed package parts: " + str(changed))
                    if mutation and not changed:
                        raise AssertionError("Requested edit was not written")
                row["status"] = "passed"
            except Exception as error:
                row.update(status="failed", error=f"{type(error).__name__}: {error}")
            finally:
                if digest(source) != original_hash:
                    raise AssertionError("Original input changed")
            if samples:
                row.update(first_ms=samples[0], warm_median_ms=statistics.median(samples[1:]) if len(samples)>1 else None,
                           p95_ms=sorted(samples)[min(len(samples)-1, int(len(samples)*0.95))], samples_ms=samples)
            results.append(row)
    if args.go_report:
        for build in json.loads(args.go_report.read_text(encoding="utf-8")):
            source, output = Path(build["input"]), Path(build["output"])
            row = {"engine": "Go", "input": str(source.resolve()), "output": str(output.resolve()), "operation": build["operation"]}
            try:
                with WordFile(source) as document:
                    expected = document.vba_modules()
                if build["operation"] == "module-edit":
                    target = sorted(expected)[0]
                    expected[target] = expected[target].rstrip("\r\n") + "\r\n' WordUp writer comparison\r\n"
                with WordFile(output) as document:
                    normalize = lambda values: {key: value.replace("\r\n", "\n").rstrip("\n") for key, value in values.items()}
                    if normalize(document.vba_modules()) != normalize(expected):
                        raise AssertionError("Go output source differs from intended edit")
                before, after = parts(source), parts(output)
                changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
                before_streams, after_streams = streams(source), streams(output)
                row.update(changed_parts=changed, changed_streams=sorted(name for name in set(before_streams) | set(after_streams)
                           if before_streams.get(name) != after_streams.get(name)))
                allowed = {"word/vbaProject.bin"} if build["operation"] == "module-edit" else set()
                if set(changed) - allowed:
                    raise AssertionError("Go output changed unrelated package parts")
                row["status"] = "passed"
            except Exception as error:
                row.update(status="failed", error=f"{type(error).__name__}: {error}")
            results.append(row)
    report = {"upstream_revision": PIN, "import_ms": import_ms, "native_word_executed": False, "results": results}
    (args.output / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report))
    return int(any(row["status"] != "passed" for row in results))


if __name__ == "__main__":
    raise SystemExit(main())
