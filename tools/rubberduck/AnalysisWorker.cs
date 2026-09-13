// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Rubberduck.VBEditor.SafeComWrappers;

namespace WordUp.Analysis {
    internal static class AnalysisWorker {
        const int MaxRequest = 4 * 1024 * 1024;
        internal sealed class Module {
            public string Name { get; set; }
            public string Kind { get; set; }
            public string Path { get; set; }
            public string Source { get; set; }
        }
        internal sealed class Request {
            public Module[] Modules { get; set; }
            public string[] Inspections { get; set; }
        }
        internal static object Analyze(string input) {
            if (input == null || input.Length > MaxRequest) throw new ArgumentException("Analysis request too large");
            var request = new JavaScriptSerializer { MaxJsonLength = MaxRequest }.Deserialize<Request>(input);
            if (request == null || request.Modules == null || request.Modules.Length == 0 || request.Modules.Length > 512)
                throw new ArgumentException("Supply 1..512 modules");
            if (request.Inspections == null || request.Inspections.Length == 0 || request.Inspections.Length > 2)
                throw new ArgumentException("Supply 1..2 integrated inspections");
            var editor = new SnapshotEditor("workspace", "Workspace", request.Modules.Select(module => {
                if (module == null || module.Source == null || String.IsNullOrWhiteSpace(module.Name)) throw new ArgumentException("Incomplete module");
                var kind = module.Kind == "standard" ? ComponentType.StandardModule : module.Kind == "class" ? ComponentType.ClassModule :
                    throw new NotSupportedException("Module kind is not integrated: " + module.Kind);
                return (module.Name, kind, new SourceSnapshot(module.Path, Encoding.UTF8.GetBytes(module.Source)));
            }));
            using (var state = SnapshotAnalysis.ResolveSourceOnly(editor,
                new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}}, CancellationToken.None)) {
                if (state.ModuleExceptions.Any()) throw new ArgumentException("Source failed parsing; use vba.parse for syntax locations");
                var diagnostics = request.Inspections.Distinct().SelectMany(name => SnapshotInspections.Diagnose(name,state,editor,CancellationToken.None)).ToArray();
                return new { engine="Rubberduck", scope="source-only", external_references_resolved=false,
                    native_compilation=false, diagnostics, document_modified=false };
            }
        }
        static int Main(string[] args) {
            Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal;
            if (args.Length != 1 || args[0] != "serve") return 2;
            var json = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
            while (true) {
                var line = new StringBuilder();
                int ch;
                while ((ch = Console.Read()) != -1 && ch != '\n') {
                    if (line.Length >= MaxRequest) return 2;
                    line.Append((char)ch);
                }
                if (ch == -1 && line.Length == 0) return 0;
                object result;
                try {
                    var command = json.Deserialize<string[]>(line.ToString());
                    if (command == null || command.Length != 2 || command[0] != "analyze") throw new ArgumentException("Expected analyze request");
                    result = Analyze(command[1]);
                } catch (Exception error) { result = new { error=error.Message, type=error.GetType().Name }; }
                Console.WriteLine(json.Serialize(result));
                Console.Out.Flush();
            }
        }
    }
}
