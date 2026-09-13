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
                editor_boundary = "upstream test-only in-memory editor", cases = results
            }));
            return 0;
        } catch (Exception e) {
            Console.Error.WriteLine(e.ToString());
            return 1;
        }
    }
    static void Require(bool condition, string message) { if (!condition) throw new Exception(message); }
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
        var vbe=MockVbeBuilder.BuildFromModules("SourceProof",modules,new ReferenceLibrary[0]);
        using (var parser=MockParser.Create(vbe.Object)) {
            parser.Parse(new CancellationTokenSource());
            Require(parser.State.Status==ParserState.Ready,"Resolver did not reach Ready: "+parser.State.Status);
            var declarations=parser.State.DeclarationFinder.AllUserDeclarations.ToArray();
            verify(declarations);
            return new { name, duration_ms=timer.Elapsed.TotalMilliseconds, declarations=declarations.Length };
        }
    }
}
