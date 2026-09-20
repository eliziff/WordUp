# Working on WordUp

Use WordUp to complete the user's document or template task. Keep originals and private inputs safe; preserve content, formatting, fields and unknown features outside the intended change. Use separate working copies.

Choose the shortest useful route. Native object-model access, arbitrary VBA, package XML, rendering, comparisons and suites are available independently; there is no required workflow. Use help when unfamiliar, doctor for environment diagnosis, and selftest when an integration check is useful. A working session does not need to repeat startup checks.

Keep iteration fast: reuse warm sessions, batch native access where semantics permit, and measure unexpected latency before changing timeouts. Prefer baseline/final evidence and basic timings; detailed tracing is optional. Retain expensive captures so analysis can be repeated without rerunning Word.

Verify the intended result through the entry point users will run, with existing revisions intact when relevant. Choose proportionate assertions, expected XML or native renders. A build is not a compile, execution is not output correctness, and a passing example is not universal coverage. Report observations and unresolved gaps accurately. Investigate mismatches rather than changing expectations to obtain a pass.

Make editing recoverable: normally one undo record per user action, with application state restored on success and failure. Capture XML outside that record. Use guarded writes and artifact hashes; deployment retains its native acceptance and backup requirements. A private desktop separates UI, not OS authority.

## Repository development

- Run Go through tools/go.ps1 (Windows) or tools/go.sh (Unix), which retain the pinned toolchain and caches. WORDUP_GO can select an installed toolchain.
- This repository currently has no released users; remove obsolete internal paths rather than adding compatibility shims. This does not authorize discarding user document features.
- Keep source roots exactly package/, vba/, forms/ and assets/. Preserve unrelated work and private project material.
- Run checks relevant to the change. No benchmark campaign is required for routine work.

## References when needed

- [Commands and source format](README.md), [API](docs/API.md), [limits](docs/LIMITS.md)
- [Word behavior](docs/WORD-OPERATIONS.md), [diagnostics](docs/DIAGNOSTICS.md)
- [Expected XML](docs/EXPECTED-RESULTS.md), [corpus and output review](docs/CORPUS.md), [optional tracing](docs/SERIAL-TRACES.md)
- [Reusable template guidance](docs/skills/wordup-template-builder/SKILL.md), [component evaluation](tools/bench/README.md)
