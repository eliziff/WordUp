using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;

class Program {
    [MTAThread]
    static int Main(string[] args) {
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal;
        var json = new JavaScriptSerializer { MaxJsonLength = 4194304 };
        if(args.Length==1 && args[0]=="serve") {
            string line;
            while((line=Console.ReadLine())!=null) {
                var request=json.Deserialize<string[]>(line);
                if(request.Length==0 || request[0]=="serve") return 1;
                Main(request);
            }
            return 0;
        }
        var timer = Stopwatch.StartNew();
        try {
            object result;
            if (args.Length == 2 && args[0] == "validate") {
                using (var document = WordprocessingDocument.Open(args[1], false)) {
                    var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Take(101).ToArray();
                    result = new { engine = "Microsoft Open XML SDK 3.3.0", valid = errors.Length == 0, truncated = errors.Length > 100,
                        errors = errors.Take(100).Select(e => new { id=e.Id, description=e.Description, type=e.ErrorType.ToString(),
                            part=e.Part == null ? null : e.Part.Uri.ToString(), xpath=e.Path == null ? null : e.Path.XPath }).ToArray(),
                        document_modified=false, duration_ms=timer.Elapsed.TotalMilliseconds,
                        working_set_bytes=Process.GetCurrentProcess().WorkingSet64,
                        peak_working_set_bytes=Process.GetCurrentProcess().PeakWorkingSet64 };
                }
            } else if (args.Length == 6 && args[0] == "patterns") {
                int pid = Int32.Parse(args[1]);
                using (var automation = new UIA3Automation()) {
                    var roots = args[2].Split(',').Distinct().Select(h => automation.FromHandle(new IntPtr(Int64.Parse(h)))).ToArray();
                    foreach(var root in roots)
                        if (root.Properties.ProcessId.Value != pid) throw new Exception("Window is not owned by the requested process");
                    var lookup = Stopwatch.StartNew();
                    FlaUI.Core.Conditions.ConditionBase condition = automation.ConditionFactory.ByName(args[3]);
                    if(args[5]!="") {
                        FlaUI.Core.Definitions.ControlType type;
                        if(!Enum.TryParse(args[5],true,out type)) throw new Exception("Unknown control type");
                        condition=condition.And(automation.ConditionFactory.ByControlType(type));
                    }
                    var found = args[3] == "" ? roots : roots.SelectMany(root => root.FindAllDescendants(condition)).ToArray();
                    if(args[3]=="" && args[5]!="" && found.Any(root => !String.Equals(root.ControlType.ToString(),args[5],StringComparison.OrdinalIgnoreCase)))
                        throw new Exception("Root does not match requested control type");
                    double lookupMs=lookup.Elapsed.TotalMilliseconds;
                    if (found.Length != 1) throw new Exception("Expected exactly one control; found " + found.Length);
                    var element = found[0];
                    if (element.Properties.ProcessId.Value != pid) throw new Exception("Control ownership mismatch");
                    string action = args[4];
                    var patternTimer=Stopwatch.StartNew();
                    var expandPattern=element.Patterns.ExpandCollapse.PatternOrDefault;
                    double patternMs=patternTimer.Elapsed.TotalMilliseconds;
                    var actionTimer=Stopwatch.StartNew();
                    switch (action) {
                        case "inspect": break;
                        case "invoke": element.Patterns.Invoke.Pattern.Invoke(); break;
                        case "default_action": element.Patterns.LegacyIAccessible.Pattern.DoDefaultAction(); break;
                        case "expand":
                        case "collapse":
                            if(expandPattern==null) throw new Exception("Control does not support ExpandCollapse");
                            string wanted=action=="expand" ? "Expanded" : "Collapsed";
                            if(expandPattern.ExpandCollapseState.Value.ToString()!=wanted) {
                                if(element.ControlType==FlaUI.Core.Definitions.ControlType.ComboBox && element.Patterns.LegacyIAccessible.IsSupported)
                                    element.Patterns.LegacyIAccessible.Pattern.DoDefaultAction();
                                else if(action=="expand") expandPattern.Expand();
                                else expandPattern.Collapse();
                                if(expandPattern.ExpandCollapseState.Value.ToString()!=wanted)
                                    throw new Exception("Control did not reach requested " + wanted + " state");
                            }
                            break;
                        case "scroll_into_view": element.Patterns.ScrollItem.Pattern.ScrollIntoView(); break;
                        default: throw new Exception("Unsupported pattern action");
                    }
                    double actionMs=actionTimer.Elapsed.TotalMilliseconds;
                    result = new { engine="FlaUI UIA3 4.0.0", pid=pid, name=element.Name, action=action,
                        expand_collapse=element.Patterns.ExpandCollapse.IsSupported,
                        expand_state=expandPattern==null ? null : expandPattern.ExpandCollapseState.Value.ToString(),
                        scroll_item=element.Patterns.ScrollItem.IsSupported,
                        value=element.Patterns.Value.IsSupported,
                        text=element.Patterns.Text.IsSupported,
                        toggle=element.Patterns.Toggle.IsSupported,
                        lookup_ms=lookupMs, pattern_ms=patternMs, action_ms=actionMs,
                        duration_ms=timer.Elapsed.TotalMilliseconds,
                        working_set_bytes=Process.GetCurrentProcess().WorkingSet64,
                        peak_working_set_bytes=Process.GetCurrentProcess().PeakWorkingSet64 };
                }
            } else { throw new Exception("Expected validate PATH or patterns PID HWND NAME ACTION"); }
            Console.WriteLine(json.Serialize(result)); return 0;
        } catch (Exception e) { Console.WriteLine(json.Serialize(new { error=e.Message, type=e.GetType().FullName })); return 1; }
    }
}
