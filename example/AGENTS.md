# WordUp source workspace
Use the wordup executable. No module imports, VBE typing, or Python setup.

- Edit vba/*.bas, *.cls, and *.vba as ordinary UTF-8 files; module names and VB_Name must agree.
- New .bas files become standard modules; new .cls files become classes; .vba files are UserForm code.
- forms/<name>.json describes persistent native MSForms design, in points. Existing unsupported controls remain opaque.
- package/ is the full original Open XML package, including RibbonX XML, embedded assets and native saved parts. Do not rewrite the ZIP by hand.
- Edit package XML directly or use native Word operations for document content, styles, numbering and saved building blocks.
- build performs deterministic package and binary checks; it is not a VBA compiler.
- native execution is local Microsoft Word, never an emulator. Authorize execution only for code the user intends to run. The private desktop is UI separation, NOT a security sandbox.
- Use fresh native acceptance before deployment. Preserve fixture files and assert specific observable behavior; a passing self-test does not establish all other macros work.
- Mac static checks and Windows native tests are not Mac execution evidence.
- Raw XML, native object-model calls and arbitrary VBA are the authoring surface. Unsupported serialization must error rather than substitute an approximation.
