# WordUp source workspace

Use the wordup executable to build and inspect the artifact. Keep originals and private inputs safe. Editable source and the .wordwright baseline/index can be versioned; retain reports and captures as local evidence. Merge source, not DOTM binaries.

- Edit UTF-8 VBA in vba/; module names and VB_Name must agree. .bas files are standard modules, .cls files are classes, and .vba files are UserForm code.
- Keep roots exactly package/, vba/, forms/ and assets/. forms/<name>.json stores native MSForms design in points; unsupported controls remain opaque. package/ contains the original Open XML parts, including RibbonX and saved content. Preserve unknown properties.
- Use package XML, native object-model calls or arbitrary VBA as appropriate. xml.query and xml.patch support guarded edits; xml.snapshot captures live Word XML; xml.review, xml.verify and xml.compare inspect retained output without Word.
- Reuse a warm session and batch native access where semantics permit. Measure slow operations, including setup and evidence capture. Use help, doctor or selftest when needed, not as a startup ritual.
- Normally group each editing action into one undo record and restore prior application state on success or failure. Capture XML outside the undo record.
- Check the intended result through the actual user entry point. Preserve pending revisions when relevant. Assertions, expected XML, native renders and suites are options, not a required sequence. Baseline/final captures usually suffice; detailed tracing is optional.
- Keep expensive evidence and inspect failures before rerunning. Build success is not compilation or runtime success; execution success is not output correctness. Report only what the observed evidence establishes, on the platform tested.
- Native execution requires authority to run the code. A private desktop separates UI; it is not a security sandbox. Deployment requires fresh artifact-bound native acceptance and preserves backups.

Use wordup help or rpc help to discover operations. The supplied README and docs cover source format, Word behavior, diagnostics and optional reusable-template guidance.
