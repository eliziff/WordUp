// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace WordUp.Analysis {
    internal static class SnapshotInspections {
        internal sealed class Diagnostic {
            public string Inspection, Path, SHA256, Description;
            public int Line, Column, EndLine, EndColumn;
        }

        public static Diagnostic[] Diagnose(string name, RubberduckParserState state,
            SnapshotEditor editor, CancellationToken token) {
            return Run(name, state, editor, token).Select(result => {
                var source = editor.Sources[result.QualifiedSelection.QualifiedName];
                var selection = result.QualifiedSelection.Selection;
                var exported = result.Inspection is IParseTreeInspection tree && tree.TargetKindOfCode == CodeKind.AttributesCode;
                return new Diagnostic {
                    Inspection = name, Path = source.Path, SHA256 = source.SHA256,
                    Description = result.Description,
                    Line = exported ? selection.StartLine : source.FileLine(selection.StartLine),
                    EndLine = exported ? selection.EndLine : source.FileLine(selection.EndLine),
                    Column = selection.StartColumn, EndColumn = selection.EndColumn
                };
            }).ToArray();
        }
        // Only verified, source-only inspections. Others may require host services.
        public static IInspectionResult[] Run(string name, RubberduckParserState state,
            SnapshotEditor editor, CancellationToken token) {
            if (name != "VariableNotUsedInspection" && name != "OptionExplicitInspection")
                throw new NotSupportedException("Inspection is not integrated: " + name);
            token.ThrowIfCancellationRequested();
            if (state.ModuleExceptions.Any()) throw new InvalidOperationException("Cannot inspect failed parse");
            var type = typeof(IInspection).Assembly.GetType(
                "Rubberduck.CodeAnalysis.Inspections.Concrete." + name, true);
            var inspection = (IInspection)Activator.CreateInstance(type, state);
            if (inspection is IParseTreeInspection treeInspection) {
                foreach (var module in editor.Sources.Keys) {
                    token.ThrowIfCancellationRequested();
                    treeInspection.Listener.CurrentModuleName = module;
                    ParseTreeWalker.Default.Walk(treeInspection.Listener,
                        state.GetParseTree(module, treeInspection.TargetKindOfCode));
                }
            }
            var results = inspection.GetInspectionResults(token).ToArray();
            // Upstream inspections do not consistently observe cancellation.
            // This boundary checks between phases, not during an upstream walk.
            token.ThrowIfCancellationRequested();
            return results;
        }
    }
}
