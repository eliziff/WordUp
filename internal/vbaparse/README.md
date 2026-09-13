# Rubberduck parser integration

Source: https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.Parsing/Grammar

The original grammar files and parser predicate helper are retained under
`grammar/`, including their copyright and GPL-3.0-or-later license. The Go
predicate adapter and generated parser are part of that GPL-covered adaptation.
ANTLR's Go runtime is pinned to 4.13.1 in go.mod/go.sum.

Regenerate with `generate.ps1 -AntlrJar PATH_TO_antlr-4.13.1-complete.jar`.
Java and the ANTLR generator are contributor build tools, not end-user runtime
requirements. Regeneration translates C# predicate calls to their Go equivalents,
removes C# base classes, and requires newline-terminated input to avoid an EOF
closure rejected by ANTLR 4.13. Original source coordinates are retained.

`vba.parse` is a syntax diagnostic, not Rubberduck's full semantic analysis,
native VBA compilation, reference resolution, or conditional preprocessor.
