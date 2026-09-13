// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace WordUp.Analysis {
    // Immutable exported source plus the code-pane view expected by Rubberduck.
    // Original bytes identify the file; parsing never rewrites workspace source.
    internal sealed class SourceSnapshot {
        static readonly Regex HiddenAttribute = new Regex(
            @"^[ \t]*Attribute[ \t]+(?:[\p{L}_][\p{L}\p{N}_]*\.)?VB_[A-Za-z0-9_]+[ \t]*=",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        public string Path { get; }
        public string SHA256 { get; }
        public string Exported { get; }
        public string CodePane { get; }
        readonly int[] codePaneLines;

        public SourceSnapshot(string path, byte[] bytes) {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("Source path required");
            if (bytes == null || bytes.Length > 8 * 1024 * 1024) throw new ArgumentException("Source byte budget exceeded");
            Path = path;
            using (var hash = System.Security.Cryptography.SHA256.Create())
                SHA256 = BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            var text = new UTF8Encoding(false, true).GetString(bytes);
            if (text.StartsWith("\uFEFF", StringComparison.Ordinal)) text = text.Substring(1);
            if (text.IndexOf('\0') >= 0) throw new ArgumentException("NUL in VBA source");
            Exported = text.Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = Exported.Split('\n');
            var visible = new List<string>();
            var locations = new List<int>();
            for (int i = 0; i < lines.Length; i++) {
                if (HiddenAttribute.IsMatch(lines[i])) continue;
                visible.Add(lines[i]); locations.Add(i + 1);
            }
            CodePane = String.Join("\n", visible);
            codePaneLines = locations.ToArray();
        }
        public int FileLine(int codePaneLine) {
            if (codePaneLine < 1 || codePaneLine > codePaneLines.Length)
                throw new ArgumentOutOfRangeException(nameof(codePaneLine), "No source location for this parser line");
            return codePaneLines[codePaneLine - 1];
        }
    }

    internal sealed class SnapshotSourceProvider : ISourceCodeProvider {
        readonly IReadOnlyDictionary<QualifiedModuleName, SourceSnapshot> sources;
        readonly bool attributes;
        public SnapshotSourceProvider(IReadOnlyDictionary<QualifiedModuleName, SourceSnapshot> sources, bool attributes) {
            this.sources = sources ?? throw new ArgumentNullException(nameof(sources));
            this.attributes = attributes;
        }
        public string SourceCode(QualifiedModuleName module) {
            if (!sources.TryGetValue(module, out var source))
                throw new ArgumentException("Source snapshot missing for " + module);
            return attributes ? source.Exported : source.CodePane;
        }
    }

    internal sealed class SnapshotCompilationArguments : ICompilationArgumentsProvider {
        readonly Dictionary<string, Dictionary<string, short>> projects;
        public VBAPredefinedCompilationConstants PredefinedCompilationConstants { get; }
        public SnapshotCompilationArguments(double vbeVersion, Dictionary<string, Dictionary<string, short>> projects) {
            PredefinedCompilationConstants = new VBAPredefinedCompilationConstants(vbeVersion);
            this.projects = projects;
        }
        public Dictionary<string, short> UserDefinedCompilationArguments(string projectId) {
            if (!projects.TryGetValue(projectId, out var values)) throw new ArgumentException("Compilation settings missing: " + projectId);
            return new Dictionary<string, short>(values, StringComparer.OrdinalIgnoreCase);
        }
    }

    internal static class SnapshotParser {
        // Upstream adds a diagnostic listener without removing ANTLR's console
        // listener. Keep exceptions authoritative, without duplicate stderr noise.
        sealed class QuietTokenStreamParser : VBATokenStreamParser {
            public QuietTokenStreamParser(IParsePassErrorListenerFactory errors) : base(errors, errors) { }
            protected override Antlr4.Runtime.Tree.IParseTree Parse(Antlr4.Runtime.ITokenStream tokens,
                Antlr4.Runtime.Atn.PredictionMode prediction, Antlr4.Runtime.IParserErrorListener errors) {
                var parser = new Rubberduck.Parsing.Grammar.VBAParser(tokens);
                parser.Interpreter.PredictionMode = prediction;
                parser.RemoveErrorListeners();
                parser.AddErrorListener(errors);
                return parser.startRule();
            }
        }
        public static ModuleParser Create(IReadOnlyDictionary<QualifiedModuleName, SourceSnapshot> sources,
            double vbeVersion, Dictionary<string, Dictionary<string, short>> constants) {
            var settings = new SnapshotCompilationArguments(vbeVersion, constants);
            var cache = new CompilationArgumentsCache(settings);
            cache.ReloadCompilationArguments(sources.Keys.Select(k => k.ProjectId).Distinct());
            var preErrors = new PreprocessingParseErrorListenerFactory();
            var preprocessor = new VBAPreprocessor(new VBAPreprocessorParser(preErrors, preErrors), cache);
            var errors = new MainParseErrorListenerFactory();
            var parser = new TokenStreamParserStringParserAdapterWithPreprocessing(
                new SimpleVBAModuleTokenStreamProvider(), new QuietTokenStreamParser(errors), preprocessor);
            var annotations = typeof(IAnnotation).Assembly.GetTypes()
                .Where(t => typeof(IAnnotation).IsAssignableFrom(t) && !t.IsAbstract)
                .Select(t => (IAnnotation)Activator.CreateInstance(t));
            return new ModuleParser(new SnapshotSourceProvider(sources, false),
                new SnapshotSourceProvider(sources, true), parser, new VBAParserAnnotationFactory(annotations));
        }
    }
}
