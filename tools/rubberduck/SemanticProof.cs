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
