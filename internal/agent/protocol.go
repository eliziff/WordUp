package agent

import (
	"bufio"
	"context"
	"encoding/json"
	"fmt"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"io"
	"strings"
	"time"
)

type Tool struct {
	Name        string         `json:"name"`
	Description string         `json:"description"`
	InputSchema map[string]any `json:"inputSchema"`
}

func Tools() []Tool {
	ds := [][2]string{{"example", "Create and build an integrated editable native template: output."}, {"selftest", "Generate own integration fixture and test it in actual local Word; saves assertions/failures. Requires --execute, optional output."}, {"deploy", "Stage exact native-tested artifact path, acceptance reference, destination output. Automatically waits locally for Word to close; backup and stale-target guards."}, {"activate", "Retry a durable activation plan path after an interrupted local worker. Never closes Word."}, {"help", "Read source layout, operations and native ABI help."}, {"doctor", "Detect local native Word. Does not execute or emulate it."}, {"new", "Create a native source workspace: name,output."}, {"import", "Import an actual DOCX/DOTM package: path,output."}, {"inspect", "Inventory an artifact: path."}, {"files", "List source files; limit bounds the response."}, {"read", "Read workspace path and its optimistic concurrency hash."}, {"write", "Write UTF-8 text or base64 to path; expected_sha256 protects existing edits."}, {"search", "Literal case-insensitive source search: query,limit."}, {"build", "Build a native DOTM; optional output. No VBA execution claim."}, {"check", "Check all source modules with Rubberduck VBA grammar, plus XML/Ribbon schema and callback diagnostics. Returns per-file locations and parsed/skipped counts; does not establish native compilation."}, {"compat", "Conservative conditional Windows/Mac portability review, NOT a Mac pass."}, {"reference.document", "Read native styles and paragraph/property observations from path."}, {"reference.image", "Return reference/screenshot image pixels to the vision-capable agent."}, {"image.compare", "Exact-size pixel comparison: reference,path,tolerance,optional output difference image."}, {"native.start", "Start owned local Word lazily; requires installed Word, not a VM."}, {"native.call", "Raw native operation; see help. op=profile accepts steps (run/eval/get/invoke/put), executes sequentially and returns observations with results, errors and wall-clock duration_ms, retaining partial evidence on failure. Persistent host. Macro execution requires launch --execute."}, {"native.stop", "Close ONLY the app-owned Word and private desktop."}, {"test", "Execute assertion suite against exact artifact path; fresh=true uses an independent process. Returns and saves truthful pass/fail/not_run evidence."}}
	ds = append(ds, [2]string{"structure.source", "Return an editable standalone VBA paragraph detection core, separate from house-style mapping. Read-only native evidence plus unresolved journal-marker candidates; no installation or execution."})
	ds = append(ds, [2]string{"structure.inspect", "Inspect document path without Word: source-located paragraphs, style inheritance, direct versus inherited outline levels, character-weighted direct formatting, resolved numbering definitions, table/textbox containment and raw evidence XML. Evidence adapter, not editorial truth."})
	ds = append(ds, [2]string{"structure.resolve", "Resolve generic heading candidates and hierarchy from package, style, outline, formatting and sequence evidence while retaining ambiguity: path."})
	ds = append(ds, [2]string{"structure.compare", "Compare resolved structure with model-authored silver XML: reference. Sources named in the silver are resolved relative to the workspace."})
	ds = append(ds,
		[2]string{"component.list", "List bundled editable template components and their capabilities."},
		[2]string{"component.get", "Return one bundled component by component ID, or a local component.json bundle by path, without changing the workspace."},
		[2]string{"component.add", "Vendor one bundled component ID or local component.json bundle path into the workspace with typed parameters, collision checks and modification protection."},
		[2]string{"component.status", "Report whether an installed component remains identical to its starting source: component."},
		[2]string{"component.diff", "Compare installed source with the same parameterized bundled or local starting version; never overwrites edits."},
	)
	props := map[string]any{}
	ds = append(ds, [2]string{"test.freeze", "Freeze passing native report reference into a new local output directory, including the retained artifact, declared inputs and run evidence. Returns a portable bundle.json reference accepted by test.replay and test.compare; verifies hashes on load. Existing destinations are never overwritten. XML is copied unchanged. Uncaptured external inputs are not made hermetic."})
	ds = append(ds, [2]string{"xml.verify", "Verify actual XML path against expected XML reference directly. No manifest or annotations. comparison defaults to exact; semantic is explicit namespace-aware equality. No ignored content. Mismatch returns an error plus hashes and difference locations. Does not run Word."})
	props["comparison"] = map[string]any{"type": "string", "enum": []string{"exact", "semantic"}, "description": "xml.verify comparison; defaults to exact bytes."}
	props["namespaces"] = map[string]any{"type": "object", "additionalProperties": map[string]any{"type": "string"}, "description": "XPath prefix to namespace URI bindings, independent of source prefixes."}
	ds = append(ds, [2]string{"xml.query", "Read-only XPath over UTF-8 XML path, or DOCX/DOTM path with explicit part (e.g. word/styles.xml), using query and namespaces. Returns scalar values or bounded matches with original element byte offsets and source hashes. limit=1..100 (default 20); truncated is not an exact total. Does not start Word or rewrite XML."})
	props["xml_policy"] = map[string]any{"type": "object", "description": "Explicit ignored XML attributes/elements, each a {Space: namespace URI, Local: local name} object. No automatic exclusions.", "properties": map[string]any{"attributes": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}, "elements": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}}, "additionalProperties": false}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["context_parts"] = map[string]any{"type": "array", "items": map[string]any{"type": "string"}, "description": "Flat OPC part names to prioritize for bounded difference context (up to eight); this does not filter comparison or change equality."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["generated_toc_bookmarks"] = map[string]any{"type": "boolean", "description": "Opt-in canonical names for generated _Toc<number> bookmarks, hyperlink anchors and field references, by bookmark document order. Preserves positions, content and link correspondence; raw bytes/hashes remain unchanged."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["generated_comment_ids"] = map[string]any{"type": "boolean", "description": "Opt-in mapping of modern comment paragraph and durable IDs through comment identity. Preserves reply parents, resolution state, text and timestamps; rejects unresolved links. Supply the complete Flat OPC XML including comment parts."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["attribute_equivalences"] = map[string]any{"type": "array", "description": "Explicit equivalent values for one XML element/attribute pair. Each rule has element and attribute {Space,Local}, and values:string[]. Only listed values are equivalent; other elements and values stay compared.", "items": map[string]any{"type": "object", "properties": map[string]any{"element": map[string]any{"type": "object"}, "attribute": map[string]any{"type": "object"}, "values": map[string]any{"type": "array", "items": map[string]any{"type": "string"}}}, "required": []string{"element", "attribute", "values"}, "additionalProperties": false}}
	props["compilation_constants"] = map[string]any{"type": "object", "description": "VBA conditional compilation constants; values are Boolean, string, or number. Defaults target the local 64-bit platform and VBA7."}
	ds = append(ds, [2]string{"binary.inspect", "Independent read-only VBA module and form-string extraction using bundled oletools. Returns source and source hashes; never executes macros or starts Word."})
	ds = append(ds, [2]string{"openxml.validate", "Read-only Microsoft Open XML SDK validation of path; returns errors with part and XPath. Uses the bundled Office tools helper; never starts Word."})
	ds = append(ds, [2]string{"ribbon.callbacks", "Generate VBA callback signatures from Ribbon XML path using Office RibbonX Editor's mapping; detects incompatible reuse, returns source without overwriting user code."})
	ds = append(ds, [2]string{"vba.parse", "Parse VBA source text or workspace path with Rubberduck's grammar. Returns line/column syntax diagnostics without starting Word; does not establish native compilation or resolve types."})
	ds = append(ds, [2]string{"vba.analyze", "Analyze selected workspace VBA paths together using the analysis helper. Source-only Option Explicit and unused-variable inspections with hashes and file locations. Standard/class modules only; host/form metadata and external references are not integrated. Does not establish native compilation."})
	props["paths"] = map[string]any{"type": "array", "items": map[string]any{"type": "string"}}
	ds = append(ds, [2]string{"xml.compare", "Compare reference and path XML using namespaces, retaining raw hashes, byte identity, and bounded first-difference XML with source offsets. xml_policy.context_parts selects part-level diagnostic context without narrowing equality. xml_policy may explicitly exclude named attributes/elements; the policy is returned with the evidence."})
	ds = append(ds, [2]string{"vba.immediate", "Evaluate text beginning with ? in native Word and return its typed result. Uses a separate scratch VBA project; does not inspect paused-frame locals or capture the VBE Immediate output buffer. Requires --execute."})
	ds = append(ds, [2]string{"preview", "Open a fresh native-tested path in visible Word using acceptance reference. Optional document opens a copied manuscript with the template attached. Creates a local preview copy, restores automation security, changes no persistent trust settings. Requires --execute; Windows only."})
	ds = append(ds, [2]string{"test.replay", "Rerun the recorded suite from reference. Without path, use the required hash-verified retained artifact snapshot. Explicit path tests a new candidate. Declared suite inputs are hash-verified snapshots with fresh working copies; undeclared external files are not frozen. fresh=true uses an independent hidden host. Requires --execute. Each run keeps a distinct saved_report and XML evidence."})
	ds = append(ds, [2]string{"test.compare", "Compare baseline reference and candidate path native report files from the same suite. Automatically compares hash-verified XML snapshots using optional xml_policy, alongside asserted-field parity differences, per-step timing ratios, and separate task_timings for completed asynchronous polls. Poll-call latency is not macro duration. Does not execute Word or certify untested features."})
	ds = append(ds, [2]string{"restore", "Restore the previous template from an installed activation plan path; refuses newer user edits and waits for Word to be closed."})
	ds = append(ds, [2]string{"compile", "Build, compile in real Word and automatically sign a DOTM. Creates a local non-exportable certificate once and reuses it; no certificate parameters needed. Requires --execute. Optional output; cached exact unchanged outputs return immediately."})
	ds = append(ds, [2]string{"sign", "Sign an unsigned path to a NEW output DOTM using an automatically created or reused local certificate; optional advanced signing.thumbprint, signing.store_location, signing.signtool and signing.timestamp_url. Requires --execute, Windows SDK SignTool and registered Microsoft Office SIP. Creates legacy, agile and V3 signatures and verifies the strongest V3 digest; publisher trust is reported separately. Run native acceptance on the signed output before deployment."}, [2]string{"signature.verify", "Verify the newest VBA signature and certificate trust with Microsoft Office SIP: path, optional signing.signtool. Does not execute macros."})
	props["signing"] = map[string]any{"type": "object", "properties": map[string]any{"thumbprint": map[string]any{"type": "string"}, "signtool": map[string]any{"type": "string"}, "timestamp_url": map[string]any{"type": "string"}, "store_location": map[string]any{"type": "string"}}, "additionalProperties": false}
	props["native_options"] = map[string]any{"type": "object", "description": "native.start options: memory_limit_mb (default 2048), cpu_percent (default 50), startup_timeout_ms. Effective at host creation only."}
	for _, key := range []string{"path", "output", "name", "text", "base64", "expected_sha256", "query", "reference", "document", "part"} {
		props[key] = map[string]any{"type": "string"}
	}
	props["component"] = map[string]any{"type": "string", "description": "Stable bundled component ID such as structure.detect."}
	props["parameters"] = map[string]any{"type": "object", "additionalProperties": map[string]any{"type": "string"}, "description": "Declared component adaptation values. Currently module_prefix is a validated VBA identifier; undeclared values are rejected."}
	for _, key := range []string{"timeout_ms", "tolerance", "limit"} {
		props[key] = map[string]any{"type": "integer"}
	}
	props["fresh"] = map[string]any{"type": "boolean"}
	props["operation"] = map[string]any{"type": "object", "description": NativeHelp}
	props["suite"] = map[string]any{"type": "object", "description": "schema=1,name,steps:[{name,operation,assert:[{kind,path,expected,tolerance}],timeout_ms}], optional platforms and require_compile"}
	tools := []Tool{}
	for _, d := range ds {
		tools = append(tools, Tool{d[0], d[1], map[string]any{"type": "object", "properties": props, "additionalProperties": false}})
	}
	return tools
}

type request struct {
	JSONRPC string          `json:"jsonrpc,omitempty"`
	ID      json.RawMessage `json:"id,omitempty"`
	Method  string          `json:"method"`
	Params  json.RawMessage `json:"params,omitempty"`
}

// Serve uses newline-delimited JSON. MCP negotiates the explicitly supported
// 2025 revisions; newer clients can negotiate back. We do not claim a newer
// protocol's capability set or mix non-JSON logging into stdout.
func Serve(ctx context.Context, e *Engine, in io.Reader, out io.Writer, mcp bool) error {
	defer e.Close()
	scan := bufio.NewScanner(in)
	scan.Buffer(make([]byte, 65536), 16<<20)
	enc := json.NewEncoder(out)
	initialized := false
	ready := false
	emit := func(id json.RawMessage, v any, err error, duration time.Duration) error {
		if id == nil {
			id = json.RawMessage("null")
		}
		r := map[string]any{"id": id}
		if mcp {
			r["jsonrpc"] = "2.0"
		}
		if err != nil {
			if mcp {
				r["error"] = map[string]any{"code": -32602, "message": err.Error()}
			} else {
				r["error"] = verify.ErrorValue(err)
				r["result"] = v
			}
		} else {
			r["result"] = v
		}
		if !mcp {
			r["duration_ms"] = float64(duration.Microseconds()) / 1000
		}
		return enc.Encode(r)
	}
	for scan.Scan() {
		start := time.Now()
		var q request
		if err := json.Unmarshal(scan.Bytes(), &q); err != nil {
			if err = emit(nil, nil, err, 0); err != nil {
				return err
			}
			continue
		}
		if mcp && q.JSONRPC != "2.0" {
			if err := emit(q.ID, nil, fmt.Errorf("JSON-RPC 2.0 required"), 0); err != nil {
				return err
			}
			continue
		}
		if mcp {
			switch q.Method {
			case "initialize":
				if initialized {
					if err := emit(q.ID, nil, fmt.Errorf("already initialized"), 0); err != nil {
						return err
					}
					continue
				}
				var p struct {
					Version string `json:"protocolVersion"`
				}
				if err := json.Unmarshal(q.Params, &p); err != nil {
					return err
				}
				version := p.Version
				if version != "2025-11-25" && version != "2025-06-18" && version != "2025-03-26" {
					version = "2025-11-25"
				}
				initialized = true
				if err := emit(q.ID, map[string]any{"protocolVersion": version, "serverInfo": map[string]any{"name": "wordup", "version": project.Version}, "capabilities": map[string]any{"tools": map[string]any{}}, "instructions": project.AgentInstructions}, nil, 0); err != nil {
					return err
				}
				continue
			case "notifications/initialized":
				ready = initialized
				continue
			case "ping":
				if err := emit(q.ID, map[string]any{}, nil, 0); err != nil {
					return err
				}
				continue
			}
			if strings.HasPrefix(q.Method, "notifications/") {
				continue
			}
			if !ready {
				if err := emit(q.ID, nil, fmt.Errorf("initialize and notifications/initialized are required"), 0); err != nil {
					return err
				}
				continue
			}
			if q.Method == "tools/list" {
				if err := emit(q.ID, map[string]any{"tools": Tools()}, nil, 0); err != nil {
					return err
				}
				continue
			}
			if q.Method != "tools/call" {
				if err := emit(q.ID, nil, fmt.Errorf("method not supported"), 0); err != nil {
					return err
				}
				continue
			}
			var p struct {
				Name      string          `json:"name"`
				Arguments json.RawMessage `json:"arguments"`
			}
			if err := json.Unmarshal(q.Params, &p); err != nil {
				if err = emit(q.ID, nil, err, 0); err != nil {
					return err
				}
				continue
			}
			q.Method = p.Name
			q.Params = p.Arguments
		}
		var p Parameters
		if len(q.Params) > 0 {
			if err := project.ReadJSON(q.Params, &p); err != nil {
				if err = emit(q.ID, nil, err, 0); err != nil {
					return err
				}
				continue
			}
		}
		value, err := e.Call(ctx, q.Method, p)
		if mcp {
			content := []any{}
			if im, ok := value.(map[string]any); ok && q.Method == "reference.image" && im["data"] != nil {
				content = append(content, map[string]any{"type": "image", "data": im["data"], "mimeType": im["mimeType"]})
				delete(im, "data")
			}
			payload := map[string]any{"result": value}
			if err != nil {
				payload["error"] = verify.ErrorValue(err)
			}
			text, _ := json.Marshal(payload)
			content = append(content, map[string]any{"type": "text", "text": string(text)})
			r := map[string]any{"content": content, "isError": err != nil}
			if err2 := emit(q.ID, r, nil, time.Since(start)); err2 != nil {
				return err2
			}
		} else {
			if err2 := emit(q.ID, value, err, time.Since(start)); err2 != nil {
				return err2
			}
		}
	}
	return scan.Err()
}
