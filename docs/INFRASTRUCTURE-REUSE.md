# Infrastructure reuse

WordUp will reuse established open-source implementations and retain their
licenses and source provenance. Integration status must not imply native test
coverage or full upstream product functionality.

| Project | Infrastructure role | Current status |
|---|---|---|
| [Rubberduck](https://github.com/rubberduck-vba/Rubberduck) | VBA grammar, declarations and reference resolution | Go grammar is integrated into `vba.parse` and `check`; the pinned .NET semantic adapter is callable through `vba.analyze` and the packaged helper. GPL-3.0-or-later. Standard/class modules and the verified Option Explicit/unused-variable inspections are covered; document/form metadata, broader inspection selection, transitive reference coverage and incremental invalidation remain explicit limits. |
| [ANTLR Go runtime](https://github.com/antlr4-go/antlr) | Run generated grammar in-process | Pinned 4.13.1; BSD runtime, no Java required by end users. |
| [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) | Callback generation and Ribbon authoring metadata | Integrated callback metadata/generation; retained MIT notice in internal/office/RIBBONX-LICENSE. |
| [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) | Independent schema/semantic OOXML validation | Integrated in the lazy OfficeTools helper; SDK/Framework 3.3.0. Full upstream MIT text is bundled by build.ps1. |
| [oletools](https://github.com/decalage2/oletools) | Independent OLE/VBA/UserForm inspection | `binary.inspect` calls the bundled worker and package acceptance exercises it outside the repository. The staged Python dependency inventory and per-file license hashes are checked by the Windows packager; source-delivery and reproducible-build review remains a release gate. |
| [FlaUI](https://github.com/FlaUI/FlaUI) | UI Automation patterns and provider interoperability | Integrated UIA3 4.0.0 for owned-process patterns; native tests cover Ribbon state changes and selector rejection. |
| [antchfx/xmlquery](https://github.com/antchfx/xmlquery) | Namespace-aware XPath assertions over exact XML evidence | Integrated `xml.query`: xmlquery 1.5.1 / xpath 1.3.6; source hashes, original element offsets, bounded matches and scalar predicates. No XML rewriting. |

The combined WordUp distribution incorporating the Rubberduck grammar is
GPL-3.0-or-later. `LICENSE` contains GPLv3; `LICENSE-MIT-ORIGINAL` preserves the
notice for pre-existing MIT code. Earlier releases retain their original terms.
Convey corresponding source, regeneration inputs and notices with releases.
Generated user templates do not acquire a software license merely because this
tool processed them; separately included third-party code retains its own terms.

Initial local parser measurement on the feature-lab VBA module: 2.4 s for the
first parse in a process, 9.5 ms for its next parse using ANTLR's warmed DFA.
A deliberately incomplete `Dim x As` returned line 2, column 9 in 1 ms. These
are syntax-only measurements, not Word compile or runtime tests.
