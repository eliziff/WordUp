using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Web.Script.Serialization;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

// Contributor proof, not a Word emulator or a production analysis endpoint.
class SemanticProof {
    static int Main() {
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal;
        try {
            CheckSourceMapping();
            var modules = new[] {
                ("Caller", "Option Explicit\nPublic Sub Run()\n Dim localValue As Long\n localValue = Callee.AddOne(4)\nEnd Sub\n", ComponentType.StandardModule),
                ("Callee", "Option Explicit\nPublic Function AddOne(ByVal value As Long) As Long\n AddOne = value + 1\nEnd Function\n", ComponentType.StandardModule)
            };
            var results = new List<object>();
            results.Add(CheckInspection());
            results.Add(CheckWorkerProtocol());
            var wordLibrary = Environment.GetEnvironmentVariable("WORDUP_TEST_WORD_TYPELIB");
            if (!String.IsNullOrEmpty(wordLibrary)) results.Add(CheckWordLibrary(wordLibrary));
            results.Add(CheckSnapshotEditor(modules));
            Action<Declaration[]> crossModule = declarations => {
                Require(declarations.Single(d => d.IdentifierName == "AddOne").References.Any(r => r.QualifiedModuleName.ComponentName == "Caller"), "Cross-module invocation unresolved");
                Require(declarations.Any(d => d.IdentifierName == "localValue") && declarations.Any(d => d.IdentifierName == "value"), "Local/parameter missing");
            };
            results.Add(Check("cross-module cold", modules, crossModule));
            results.Add(Check("cross-module warm", modules, crossModule));
            var broken = modules.ToArray();
            broken[0] = ("Caller", modules[0].Item2.Replace("Callee.AddOne", "Callee.Missing"), ComponentType.StandardModule);
            results.Add(Check("wrong member is not a valid reference", broken, declarations => {
                Require(!declarations.Single(d => d.IdentifierName == "AddOne").References.Any(r => r.QualifiedModuleName.ComponentName == "Caller"), "Wrong member falsely resolved");
            }));
            results.Add(Check("conditional declarations", new[] {
                ("Conditional", "#Const Enabled = True\n#If Enabled Then\nPublic Sub Chosen()\nEnd Sub\n#Else\nPublic Sub Rejected()\nEnd Sub\n#End If\n", ComponentType.StandardModule)
            }, declarations => {
                Require(declarations.Any(d => d.IdentifierName == "Chosen") && !declarations.Any(d => d.IdentifierName == "Rejected"), "Inactive branch leaked declarations");
            }));
            results.Add(Check("property accessors and With binding", new[] {
                ("Box", "Option Explicit\nPrivate stored As Long\nPublic Property Get Value() As Long\n Value = stored\nEnd Property\nPublic Property Let Value(ByVal rhs As Long)\n stored = rhs\nEnd Property\n", ComponentType.ClassModule),
                ("UseBox", "Option Explicit\nPublic Sub Run()\n Dim box As New Box\n Dim found As Long\n With box\n .Value = 4\n found = .Value\n End With\nEnd Sub\n", ComponentType.StandardModule)
            }, declarations => {
                var accessors=declarations.Where(d => d.IdentifierName == "Value").ToArray();
                Require(accessors.Length==2, "Property get/let were collapsed");
                Require(accessors.All(d => d.References.Any(r => r.QualifiedModuleName.ComponentName == "UseBox")), "With property references unresolved");
            }));
            Console.WriteLine(new JavaScriptSerializer().Serialize(new {
                engine = "Rubberduck fae50adab188126a5e7d2a1cefc3328cc18af482",
                status = "passed", native_word_executed = false,
                editor_boundaries = new[]{"read-only workspace snapshots", "upstream test-only in-memory editor"}, cases = results
            }));
            return 0;
        } catch (Exception e) {
            Console.Error.WriteLine(e.ToString());
            return 1;
        }
    }
    static void Require(bool condition, string message) { if (!condition) throw new Exception(message); }
    static object CheckWorkerProtocol() {
        var timer = Stopwatch.StartNew();
        var json = new JavaScriptSerializer();
        Func<string,string> request = kind => json.Serialize(new {
            Modules = new[]{new {Name="Example",Kind=kind,Path="vba/Example.bas",Source="Public Sub Run()\nEnd Sub\n"}},
            Inspections = new[]{"OptionExplicitInspection"}
        });
        var lines = new System.Collections.Concurrent.ConcurrentQueue<string>();
        var errors = new System.Collections.Concurrent.ConcurrentQueue<string>();
        using (var worker = new Process { StartInfo = new ProcessStartInfo {
            FileName=System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"WordUp.Analysis.exe"),
            Arguments="serve",UseShellExecute=false,CreateNoWindow=true,
            RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true
        }}) {
            worker.OutputDataReceived += (s,e) => { if(e.Data != null) lines.Enqueue(e.Data); };
            worker.ErrorDataReceived += (s,e) => { if(e.Data != null) errors.Enqueue(e.Data); };
            worker.Start();
            worker.BeginOutputReadLine(); worker.BeginErrorReadLine();
            worker.StandardInput.WriteLine("null");
            worker.StandardInput.WriteLine(json.Serialize(new[]{"analyze","{}"}));
            worker.StandardInput.WriteLine(json.Serialize(new[]{"analyze",request("form")}));
            worker.StandardInput.WriteLine(json.Serialize(new[]{"analyze",request("standard")}));
            worker.StandardInput.Close();
            if (!worker.WaitForExit(30000)) {
                worker.Kill(); worker.WaitForExit();
                throw new Exception("Analysis worker exceeded proof deadline");
            }
            worker.WaitForExit(); // Drain asynchronous output events.
            Require(worker.ExitCode == 0 && errors.IsEmpty,"Analysis worker process failed: " + String.Join("\n",errors));
        }
        var replies = lines.Select(line => json.Deserialize<Dictionary<string,object>>(line)).ToArray();
        Require(replies.Length == 4,"Analysis worker lost protocol responses");
        Require(replies.Take(3).All(reply => reply.ContainsKey("error")),"Invalid or unsupported request was accepted");
        Require(!replies[3].ContainsKey("error") && (string)replies[3]["scope"] == "source-only" && !(bool)replies[3]["native_compilation"],
            "Worker failed recovery or overstated verification");
        var findings = (System.Collections.ArrayList)replies[3]["diagnostics"];
        Require(findings.Count == 1,"Worker did not execute inspection after errors");
        return new {name="analysis worker protocol errors and recovery",duration_ms=timer.Elapsed.TotalMilliseconds};
    }
    static object CheckInspection() {
        var timer = Stopwatch.StartNew();
        // Concrete inspections are internal upstream; instantiate the pinned
        // implementation through its public IInspection contract, without mocks.
        foreach (var used in new[]{false, true, false}) {
            var source = "Option Explicit\nPublic Function Answer() As Long\n Dim spare As Long\n Answer = " + (used ? "spare" : "42") + "\nEnd Function\n";
            var editor = new WordUp.Analysis.SnapshotEditor("workspace", "InspectionProof", new[]{
                ("Example", ComponentType.StandardModule, new WordUp.Analysis.SourceSnapshot("vba/Example.bas", System.Text.Encoding.UTF8.GetBytes(source)))
            });
            using (var state = WordUp.Analysis.SnapshotAnalysis.ResolveSourceOnly(editor,
                new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}},CancellationToken.None)) {
                Require(!state.ModuleExceptions.Any(), "Inspection fixture failed to parse");
                var findings = WordUp.Analysis.SnapshotInspections.Run("VariableNotUsedInspection", state, editor, CancellationToken.None);
                Require(findings.Length == (used ? 0 : 1), "Unused variable inspection missed defect or retained stale result");
                if (!used) {
                    Require(findings[0].Target.IdentifierName == "spare", "Inspection targeted wrong variable");
                    Require(findings[0].QualifiedSelection.QualifiedName.ComponentName == "Example", "Inspection targeted wrong module");
                    Require(findings[0].QualifiedSelection.Selection.StartLine == 3, "Inspection location incorrect");
                    Require(!String.IsNullOrEmpty(findings[0].Description), "Inspection resource description missing");
                }
            }
        }
        foreach (var explicitOption in new[]{false,true}) {
            var source = "Attribute VB_Name = \"Example\"\n" + (explicitOption ? "Option Explicit\n" : "") + "Public Sub Run()\nEnd Sub\n";
            var snapshot = new WordUp.Analysis.SourceSnapshot("vba/Example.bas", System.Text.Encoding.UTF8.GetBytes(source));
            var editor = new WordUp.Analysis.SnapshotEditor("workspace", "InspectionProof", new[]{("Example",ComponentType.StandardModule,snapshot)});
            using (var state = WordUp.Analysis.SnapshotAnalysis.ResolveSourceOnly(editor,
                new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}},CancellationToken.None)) {
                var findings = WordUp.Analysis.SnapshotInspections.Run("OptionExplicitInspection",state,editor,CancellationToken.None);
                Require(findings.Length == (explicitOption ? 0 : 1), "Option Explicit inspection missed defect or correction");
                if (!explicitOption) Require(snapshot.FileLine(findings[0].QualifiedSelection.Selection.StartLine) == 2, "Exported inspection location lost hidden attribute offset");
                var mapped = WordUp.Analysis.SnapshotInspections.Diagnose("OptionExplicitInspection",state,editor,CancellationToken.None);
                Require(mapped.Length == findings.Length, "Mapped inspection lost results");
                if (!explicitOption) Require(mapped[0].Path == snapshot.Path && mapped[0].SHA256 == snapshot.SHA256 && mapped[0].Line == 2 && mapped[0].Column > 0,
                    "Mapped diagnostic lost source identity or exported location");
                try {
                    WordUp.Analysis.SnapshotInspections.Run("OptionExplicitInspection",state,editor,new CancellationToken(true));
                    throw new Exception("Cancelled inspection executed");
                } catch (OperationCanceledException) { }
            }
        }
        return new {name="upstream variable and Option Explicit inspections; mapped sources and cancellation",duration_ms=timer.Elapsed.TotalMilliseconds};
    }
    static object CheckWordLibrary(string path) {
        var timer = Stopwatch.StartNew();
        var provider = new Rubberduck.Parsing.ComReflection.ComLibraryProvider();
        var library = provider.LoadTypeLibrary(path);
        Require(library != null,"Test Word type library unavailable");
        Rubberduck.VBEditor.ReferenceInfo reference;
        try { reference = provider.GetReferenceInfo(library,"Word",path); }
        finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(library); }
        var references = new WordUp.Analysis.SnapshotReferences("workspace", new[]{reference});
        Require(references.Unavailable.Count == 0,"Word reference unavailable: " + String.Join(",",references.Unavailable.Select(r => r.Error)));
        var editor = new WordUp.Analysis.SnapshotEditor("workspace","SourceProof",new[]{
            ("UseWord",ComponentType.StandardModule,new WordUp.Analysis.SourceSnapshot("vba/UseWord.bas",System.Text.Encoding.UTF8.GetBytes(
                "Option Explicit\nPublic Sub Inspect(ByVal doc As Word.Document)\n Dim title As String\n title = doc.Name\nEnd Sub\n")))
        });
        for (int run=0;run<2;run++) using (var state = WordUp.Analysis.SnapshotAnalysis.Resolve(editor,
            new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}},references,CancellationToken.None)) {
            Require(!state.ModuleExceptions.Any(),"Word reference test syntax failure");
            var parameter = state.DeclarationFinder.AllUserDeclarations.Single(d => d.IdentifierName == "doc");
            Require(parameter.AsTypeDeclaration != null && parameter.AsTypeDeclaration.IdentifierName == "Document","Word.Document type unresolved");
            Require(state.DeclarationFinder.AllDeclarations.Where(d => d.IdentifierName == "Name")
                .Sum(d => d.References.Count(r => r.QualifiedModuleName.ComponentName == "UseWord")) == 1,
                "Document.Name member unresolved or stale references accumulated");
        }
        var invalid = new WordUp.Analysis.SnapshotReferences("workspace",new[]{new Rubberduck.VBEditor.ReferenceInfo(Guid.Empty,"Word",path,0,0)});
        Require(invalid.Unavailable.Count == 1 && invalid.DeclarationCount == 0 && invalid.ProjectReferences.Count == 0,
            "Mismatched reference falsely reported loaded");
        var missing = new WordUp.Analysis.SnapshotReferences("workspace",new[]{new Rubberduck.VBEditor.ReferenceInfo(reference.Guid,"Word",
            System.IO.Path.Combine(System.IO.Path.GetTempPath(),Guid.NewGuid().ToString("N"),"missing.olb"),reference.Major,reference.Minor)});
        Require(missing.Unavailable.Count == 1 && missing.DeclarationCount == 0,"Missing reference falsely reported loaded");
        return new {name="installed Word type library and identity rejection",duration_ms=timer.Elapsed.TotalMilliseconds,
            declarations=references.DeclarationCount,hashes=references.Hashes};
    }
    static object CheckSnapshotEditor((string,string,ComponentType)[] modules) {
        var timer = Stopwatch.StartNew();
        var editor = new WordUp.Analysis.SnapshotEditor("workspace", "SourceProof", modules.Select(m =>
            (m.Item1, m.Item3, new WordUp.Analysis.SourceSnapshot("vba/" + m.Item1 + ".bas", System.Text.Encoding.UTF8.GetBytes(m.Item2)))));
        using (var state = WordUp.Analysis.SnapshotAnalysis.ResolveSourceOnly(editor,
            new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}},CancellationToken.None)) {
            Require(!state.ModuleExceptions.Any(),"Snapshot parsing failed");
            Require(state.DeclarationFinder.AllUserDeclarations.Single(d => d.IdentifierName == "AddOne")
                .References.Any(r => r.QualifiedModuleName.ComponentName == "Caller"),"Snapshot cross-module invocation unresolved");
            try { editor.Vbe.ActiveVBProject = editor.Project("workspace"); throw new Exception("Snapshot permitted mutation"); }
            catch (NotSupportedException) { }
            return new { name="workspace snapshot semantic resolution (no mocks)", duration_ms=timer.Elapsed.TotalMilliseconds };
        }
    }
    static void CheckSourceMapping() {
        var original = "\uFEFFAttribute VB_Name = \"Example\"\r\nOption Explicit\r\nPublic Sub Run()\r\nAttribute Run.VB_Description = \"entry\"\r\n 'Attribute VB_Name = \"comment\"\r\n Dim message As String\r\n message = \"Attribute VB_Name = text\"\r\nEnd Sub\r\n";
        var source = new WordUp.Analysis.SourceSnapshot("vba/Example.bas", System.Text.Encoding.UTF8.GetBytes(original));
        Require(source.FileLine(1)==2 && source.FileLine(3)==5 && source.FileLine(5)==7, "Code-pane source mapping lost hidden attribute offsets");
        Require(source.CodePane.Contains("'Attribute VB_Name") && source.CodePane.Contains("message = \"Attribute"), "Comment/string incorrectly removed");
        Require(source.Exported.Contains("Attribute Run.VB_Description"), "Attributes pass lost member metadata");
        var key = new Rubberduck.VBEditor.QualifiedModuleName("Project", "", "Example", "snapshot");
        var sources = new Dictionary<Rubberduck.VBEditor.QualifiedModuleName, WordUp.Analysis.SourceSnapshot>{{key,source}};
        Require(new WordUp.Analysis.SnapshotSourceProvider(sources,false).SourceCode(key)==source.CodePane, "Code-pane provider mismatch");
        Require(new WordUp.Analysis.SnapshotSourceProvider(sources,true).SourceCode(key)==source.Exported, "Attributes provider mismatch");
        var constants = new Dictionary<string, Dictionary<string, short>>{{"snapshot",new Dictionary<string,short>()}};
        var parsed = WordUp.Analysis.SnapshotParser.Create(sources,7.1,constants).Parse(key,CancellationToken.None);
        Require(parsed.CodePaneParseTree != null && parsed.AttributesParseTree != null && parsed.Attributes.Count > 0,
            "Real module parser lost code-pane or attribute results");
        sources[key] = new WordUp.Analysis.SourceSnapshot("vba/Example.bas",System.Text.Encoding.UTF8.GetBytes(
            "Attribute VB_Name = \"Example\"\nOption Explicit\nPublic Sub Run()\nAttribute Run.VB_Description = \"entry\"\nDim As Long\nEnd Sub\n"));
        try {
            WordUp.Analysis.SnapshotParser.Create(sources,7.1,constants).Parse(key,CancellationToken.None);
            throw new Exception("Invalid workspace code accepted");
        } catch (Rubberduck.Parsing.VBA.Parsing.ParsingExceptions.SyntaxErrorException error) {
            Require(sources[key].FileLine(error.LineNumber)==5,"Parser diagnostic did not map to the editable file line");
        }
    }
    static object Check(string name, (string,string,ComponentType)[] modules, Action<Declaration[]> verify) {
        var timer=Stopwatch.StartNew();
        string[] snapshotSignatures;
        var editor = new WordUp.Analysis.SnapshotEditor("workspace", "SourceProof", modules.Select(m =>
            (m.Item1,m.Item3,new WordUp.Analysis.SourceSnapshot("vba/" + m.Item1 + ".bas",System.Text.Encoding.UTF8.GetBytes(m.Item2)))));
        using (var snapshot = WordUp.Analysis.SnapshotAnalysis.ResolveSourceOnly(editor,
            new Dictionary<string,Dictionary<string,short>>{{"workspace",new Dictionary<string,short>()}},CancellationToken.None)) {
            Require(!snapshot.ModuleExceptions.Any(),"Snapshot parse errors");
            verify(snapshot.DeclarationFinder.AllUserDeclarations.ToArray());
            snapshotSignatures = Signatures(snapshot.DeclarationFinder.AllUserDeclarations);
        }
        var snapshotMs = timer.Elapsed.TotalMilliseconds;
        timer.Restart();
        var vbe=MockVbeBuilder.BuildFromModules("SourceProof",modules,new ReferenceLibrary[0]);
        using (var parser=MockParser.Create(vbe.Object)) {
            parser.Parse(new CancellationTokenSource());
            Require(parser.State.Status==ParserState.Ready,"Resolver did not reach Ready: "+parser.State.Status);
            var declarations=parser.State.DeclarationFinder.AllUserDeclarations.ToArray();
            verify(declarations);
            Require(snapshotSignatures.SequenceEqual(Signatures(declarations)), "Snapshot/upstream declaration or reference mismatch: " + name);
            return new { name, snapshot_ms=snapshotMs, upstream_test_editor_ms=timer.Elapsed.TotalMilliseconds, declarations=declarations.Length };
        }
    }
    static string[] Signatures(IEnumerable<Declaration> declarations) => declarations.Select(d =>
        d.QualifiedModuleName.ComponentName + "|" + d.IdentifierName + "|" + d.DeclarationType + "|" + d.AsTypeName + "|" +
        String.Join(",",d.References.Select(r => r.QualifiedModuleName.ComponentName + ":" + r.Selection).OrderBy(s => s,StringComparer.Ordinal)))
        .OrderBy(s => s,StringComparer.Ordinal).ToArray();
}
