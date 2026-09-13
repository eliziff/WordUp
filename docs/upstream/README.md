# Offline upstream documentation

This directory holds unmodified upstream documentation and license files.
`sources.json` records repository URLs, exact commits, archive hashes and every
copied file's SHA-256. Attribution belongs to the upstream authors and respective
projects. Their licenses apply to these files independently of WordUp's license.
Microsoft VBA documentation uses CC BY 4.0; its code samples have a separate MIT
license. Both notices are retained in `vba-docs`.

The VBA snapshot includes the entire Word conceptual section and Word API files,
the shared Office API/concepts, VBA language and Forms reference, and library
reference. It excludes other applications' API references. It is a snapshot of
the published source, not a claim that Microsoft documents every Word behavior.

Other snapshots retain repository Markdown/reStructuredText documentation,
documentation images and license notices for Open XML SDK, FlaUI, oletools,
Rubberduck, RibbonX Editor, ANTLR Go and pyOpenVBA. External wikis and separately
hosted API sites are not mirrored. Open XML SDK, FlaUI, oletools, Rubberduck and
ANTLR snapshots target the versions used by WordUp; other reference repositories
are pinned at the commit recorded in the manifest.

Search locally, with no index or background process:

```powershell
rg -n -i 'wdUndefined|ScreenUpdating|UndoRecord' docs/upstream/vba-docs
rg -n -i 'ExpandCollapse|InvokePattern' docs/upstream/flaui
rg -n -i 'validation' docs/upstream/open-xml-sdk
```

Maintainers can run `python tools/sync-docs.py` to restore locked snapshots,
`python tools/sync-docs.py --check` to check file hashes offline, and
`python tools/sync-docs.py --update` to resolve the configured refs again.
Use `--refresh` to reapply documentation selection from the locked commits
without upgrading upstream versions.
Python is only a maintenance tool; WordUp does not start it to use documentation.
Refreshes retain original upstream paths/content. Relative links to excluded
sections and Microsoft Learn include syntax may require the upstream website.

See [WordUp's observed Word behavior](../WORD-OPERATIONS.md) for evidence that
adds to the published reference. The original source URL for any bundled file
is `<repository>/blob/<commit>/<relative path>` using its manifest entry.
