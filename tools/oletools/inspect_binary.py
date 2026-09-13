"""Independent, read-only OLE/VBA inspection using oletools (BSD/MIT)."""
import ctypes
import hashlib
import json
import logging
import sys
import time

ctypes.windll.kernel32.SetPriorityClass(ctypes.windll.kernel32.GetCurrentProcess(), 0x4000)
logging.disable(logging.CRITICAL)
from oletools.olevba import VBA_Parser

class MemoryCounters(ctypes.Structure):
    _fields_ = [("cb", ctypes.c_ulong), ("faults", ctypes.c_ulong)] + [
        (name, ctypes.c_size_t) for name in (
            "peak_working_set", "working_set", "peak_paged_pool", "paged_pool",
            "peak_nonpaged_pool", "nonpaged_pool", "pagefile", "peak_pagefile")]

def memory_usage():
    counters = MemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    if not ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.c_void_p(-1), ctypes.byref(counters), counters.cb):
        return None
    return {"working_set_bytes": counters.working_set,
            "peak_working_set_bytes": counters.peak_working_set}

def respond(path):
    started = time.perf_counter()
    try:
    
        parser = VBA_Parser(path)
        try:
            modules = []
            for container, stream, name, source in parser.extract_macros():
                raw = source.encode("utf-8") if isinstance(source, str) else source
                modules.append({"name": name, "stream": stream,
                                "source": raw.decode("utf-8", errors="replace"),
                                "source_sha256": hashlib.sha256(raw).hexdigest()})
            forms = list(parser.extract_form_strings())
            print(json.dumps({"engine": "oletools 0.60.2", "modules": modules,
                              "memory": memory_usage(),
                              "form_strings": forms, "document_modified": False,
                              "vba_compiled": False,
                              "duration_ms": (time.perf_counter()-started)*1000}, default=str))
        finally:
            parser.close()
    except Exception as exc:
        print(json.dumps({"error": str(exc), "type": type(exc).__name__}))
        return
    
if sys.argv[1:] == ["serve"]:
    for line in sys.stdin:
        respond(json.loads(line)[0])
        sys.stdout.flush()
else:
    respond(sys.argv[1])