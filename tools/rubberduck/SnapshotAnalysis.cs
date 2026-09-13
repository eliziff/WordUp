// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.DeclarationResolving;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.ReferenceManagement;

namespace WordUp.Analysis {
    internal static class SnapshotAnalysis {
        // This entry point deliberately accepts only source-only projects. Callers
        // with external references must load declarations and priority maps first;
        // an empty reference set is not evidence that Word/MSForms calls resolved.
        public static RubberduckParserState ResolveSourceOnly(SnapshotEditor editor,
            Dictionary<string, Dictionary<string, short>> constants, CancellationToken token) => Resolve(editor,constants,null,token);
        public static RubberduckParserState Resolve(SnapshotEditor editor,
            Dictionary<string, Dictionary<string, short>> constants, SnapshotReferences references, CancellationToken token) {
            var state = new RubberduckParserState(editor.Vbe, editor, new DeclarationFinderFactory(), editor.Events);
            try {
                var manager = new SynchronousParserStateManager(state);
                var keys = editor.Sources.Keys.ToArray();
                var parser = SnapshotParser.Create(editor.Sources, 7.1, constants);
                new SynchronousParseRunner(state, manager, parser).ParseModules(keys, token);
                if (state.ModuleExceptions.Any()) return state;
                if (references != null) foreach (var declaration in references.CreateDeclarations()) state.AddDeclaration(declaration);
                state.RefreshDeclarationFinder();
                var declarations = new SynchronousDeclarationResolveRunner(state, manager,
                    (IProjectReferencesProvider)references ?? new SourceOnlyReferences());
                declarations.CreateProjectDeclarations(keys.Select(k => k.ProjectId).Distinct().ToArray());
                state.RefreshDeclarationFinder();
                declarations.RefreshProjectReferences();
                declarations.ResolveDeclarations(keys, token);
                state.RefreshDeclarationFinder();
                var links = new ModuleToModuleReferenceManager();
                new SynchronousReferenceResolveRunner(state, manager, links,
                    new SynchronousReferenceRemover(state, links), new SourceOnlyReferences()).ResolveReferences(keys, token);
                state.RefreshDeclarationFinder();
                if (keys.Any(k => state.GetModuleState(k) == ParserState.ResolverError))
                    throw new InvalidOperationException("Rubberduck source resolver failed");
                return state;
            } catch { state.Dispose(); throw; }
        }
        sealed class SourceOnlyReferences : IProjectReferencesProvider, IDocumentModuleSuperTypeNamesProvider {
            public IReadOnlyCollection<ReferencePriorityMap> ProjectReferences => new ReferencePriorityMap[0];
            public IEnumerable<string> GetSuperTypeNamesFor(DocumentModuleDeclaration document) =>
                throw new NotSupportedException("Document modules require host type-library metadata");
        }
    }
}
