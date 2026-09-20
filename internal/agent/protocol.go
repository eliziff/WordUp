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

// One descriptor per operation: ! marks a required input. Only relevant
// properties are advertised, not the entire engine parameter union per tool.
var toolSpecs = []struct{ name, parameters, description string }{
	{"styles.apply", "path! expected_sha256! recipe!", "Atomically apply a style or numbering recipe to one workspace XML source part, preserving unknown XML and inherited namespaces. Supports named styles, paragraph/run properties and multilevel numbering; no Word startup or layout claims. Requires the current full-file hash."},
	{"example", "output!", "Create and build an integrated editable native template: output."},
	{"selftest", "output fresh", "Generate own integration fixture and test it in actual local Word; saves assertions/failures. Requires --execute, optional output."},
	{"deploy", "path! reference! output!", "Stage exact native-tested artifact path, acceptance reference, destination output. Automatically waits locally for Word to close; backup and stale-target guards."},
	{"activate", "path!", "Retry a durable activation plan path after an interrupted local worker. Never closes Word."},
	{"help", "", "Read source layout, operations and native ABI help."},
	{"doctor", "", "Detect local native Word. Does not execute or emulate it."},
	{"new", "name! output!", "Create a native source workspace: name,output."},
	{"import", "path! output!", "Import an actual DOCX/DOTM package: path,output."},
	{"inspect", "path!", "Inventory an artifact: path."},
	{"files", "limit", "List source files; limit bounds the response."},
	{"read", "path! offset limit", "Read a workspace file with its full-file SHA-256; optional offset and limit retrieve a bounded byte range for large reports or evidence."},
	{"write", "path! text base64 expected_sha256", "Write UTF-8 text or base64 to path; expected_sha256 protects existing edits."},
	{"search", "query! limit", "Literal case-insensitive source search: query,limit."},
	{"build", "output", "Build a native DOTM; optional output. No VBA execution claim."},
	{"check", "compilation_constants", "Check source syntax, XML/Ribbon wiring, and emit a compact inventory of public procedures, form events, Ribbon callbacks, hotkey/context-menu registrations, and installed component state; does not establish native compilation."},
	{"compat", "", "Conservative conditional Windows/Mac portability review, NOT a Mac pass."},
	{"reference.document", "path!", "Read exact theme1.xml, styles.xml, numbering.xml and available fontTable.xml parts with hashes, inventory every OPC part and its content types/relationship graph, source-located section geometry, document/settings/relationship hashes, theme font resolution, paragraph/property observations, content-control, field/link/bookmark evidence, and artwork evidence with relationship targets from path."},
	{"reference.image", "path!", "Return reference/screenshot image pixels to the vision-capable agent."},
	{"image.compare", "reference! path! output tolerance", "Exact-size pixel comparison: reference,path,tolerance,optional output difference image."},
	{"native.start", "native_options", "Start owned local Word lazily; requires installed Word, not a VM."},
	{"native.call", "operation! timeout_ms", "Raw native operation; see help. op=profile accepts steps (run/eval/get/invoke/put), executes sequentially and returns observations with results, errors and wall-clock duration_ms, retaining partial evidence on failure. Persistent host. Macro execution requires launch --execute."},
	{"native.stop", "", "Close ONLY the app-owned Word and private desktop."},
	{"native.probe", "count native_options", "Launch count (default 1, max 50) fresh app-owned hidden Word sessions back to back, handshake, and close each. Reports per-launch startup timings, Word version, desktop isolation and p50/p95; any failed launch or leftover process is an error. This is the headless-Word gate, independent of the warm session."},
	{"test", "path! suite fresh", "Execute assertion suite against exact artifact path; fresh=true uses an independent process. Returns and saves truthful pass/fail/not_run evidence."},
	{"test.corpus", "path! suite! cases! output! reference", "Run an existing suite sequentially over cases [{id,inputs}] in a warm owned Word host, using fresh working copies. output is a parent directory for a new retained run. Optional reference is a previous corpus.json; reuse only matching passing hash-verified frozen evidence. Failed/incomplete cases fail the aggregate; retries retain earlier attempts. Requires --execute. Project rules and documents remain external."},
	{"structure.source", "", "Return an editable standalone VBA paragraph detection core, separate from house-style mapping. Read-only native evidence plus unresolved marker candidates; no installation or execution."},
	{"structure.inspect", "path!", "Inspect document path without Word: source-located paragraphs, style inheritance, direct versus inherited outline levels, character-weighted direct formatting, resolved numbering definitions, table/textbox containment and raw evidence XML. Evidence adapter, not editorial truth."},
	{"structure.resolve", "path!", "Resolve generic heading candidates and hierarchy from package, style, outline, formatting and sequence evidence while retaining ambiguity: path."},
	{"structure.compare", "reference! limit", "Compare resolved structure with model-authored silver XML: reference. Sources named in the silver are resolved relative to the workspace."},
	{"component.list", "", "List bundled editable template components and their capabilities."},
	{"component.get", "component path", "Return one bundled component by component ID, or a local component.json bundle by path, without changing the workspace."},
	{"component.add", "component path parameters", "Vendor one bundled component ID or local component.json bundle path into the workspace with typed parameters, collision checks, Ribbon-fragment preflight and modification protection."},
	{"component.status", "component!", "Report installed and bundled versions, source hashes, modification state and host-platform compatibility for an installed component: component."},
	{"component.diff", "component path parameters", "Compare installed source with the same parameterized bundled or local starting version using hashes, sizes and first differing byte; never overwrites edits or returns source text."},
	{"test.freeze", "reference! output!", "Freeze passing native report reference into a new local output directory, including the retained artifact, declared inputs and run evidence. Returns a portable bundle.json reference accepted by test.replay and test.compare; verifies hashes on load. Existing destinations are never overwritten. XML is copied unchanged. Uncaptured external inputs are not made hermetic."},
	{"xml.verify", "reference! path! part comparison", "Verify actual XML or a named OPC package part against an expected XML reference directly. comparison defaults to exact; semantic is explicit namespace-aware equality. No ignored content. Mismatch returns an error plus hashes and difference locations. Does not run Word."},
	{"xml.query", "path! query! namespaces part limit", "Read-only XPath over UTF-8 XML path, or DOCX/DOTM path with explicit part (e.g. word/styles.xml), using query and namespaces. Returns scalar values or bounded matches with original element byte offsets and source hashes. limit=1..100 (default 20); truncated is not an exact total. Does not start Word or rewrite XML."},
	{"xml.patch", "path! expected_sha256! patches!", "Atomically apply a nonempty patches array of {offset,length,text} using original UTF-8 byte offsets in a workspace XML source. Requires expected_sha256. Rejects overlaps, split characters and invalid XML; preserves bytes outside the ranges. Does not rewrite ZIP packages."},
	{"binary.inspect", "path!", "Independent read-only VBA module and form-string extraction using bundled oletools. Returns source and source hashes; never executes macros or starts Word."},
	{"openxml.validate", "path!", "Read-only Microsoft Open XML SDK validation of path; returns errors with part and XPath. Uses the bundled Office tools helper; never starts Word."},
	{"ribbon.callbacks", "path!", "Generate VBA callback signatures from Ribbon XML path using Office RibbonX Editor's mapping; detects incompatible reuse, returns source without overwriting user code."},
	{"vba.parse", "path text compilation_constants", "Parse VBA source text or workspace path with Rubberduck's grammar. Returns line/column syntax diagnostics without starting Word; does not establish native compilation or resolve types."},
	{"vba.analyze", "paths", "Analyze selected workspace VBA paths together using the analysis helper. Source-only Option Explicit and unused-variable inspections with hashes and file locations. Standard/class modules only; host/form metadata and external references are not integrated. Does not establish native compilation."},
	{"xml.compare", "reference! path! part xml_policy", "Compare reference and path XML, or the same named OPC package part from each, using namespaces. Retains raw hashes, byte identity, and bounded first-difference XML with source offsets. xml_policy.context_parts selects part-level diagnostic context without narrowing equality. xml_policy may explicitly exclude named attributes/elements; the policy is returned with the evidence."},
	{"xml.review", "reference! path! part offset limit", "Read-only before/after review of changed WordprocessingML paragraphs in XML/Flat OPC or an explicit package part. Paginated rows show text, revision/field/control markup and direct formatting with source hashes and byte locations. limit defaults to 20 (maximum 100). Paragraph alignment is positional within each story, not semantic. Does not approve editorial output, resolve inherited formatting or establish native serial equivalence."},
	{"alr.grade", "reference! path!", "Grade every text change between a before (reference) and after (path) Flat OPC snapshot of a manuscript against the ported ALR code map: body paragraphs and footnotes are aligned, and each change is classed as an oracle citation normalization, a formatting-only rewrite (quotes, punctuation, whitespace, heading prefix), a case-only change, an emptied paragraph, protected field text, or unexpected. Returns counts by class, oracle refusals, deleted/inserted paragraphs and bounded excerpts, unexpected first. Does not start Word."},
	{"vba.immediate", "text paused symbols timeout_ms", "Evaluate text beginning with ? in native Word. Default mode uses a separate scratch VBA project; paused:true evaluates against the owned paused VBA frame through the VBE Immediate Window, and optional symbols returns best-effort values with each unavailable symbol reported separately. Requires --execute."},
	{"preview", "path! reference document", "Open an isolated copy in visible Word. Attaches the template, enables automatic style updates, applies styles, saves and verifies attachment state. Optional document supplies a manuscript; otherwise creates a new document. Optional acceptance reference is reported, not required or treated as deployment approval. Requires --execute; Windows only. No persistent trust changes."},
	{"test.replay", "reference! path fresh", "Rerun the recorded suite from reference. Without path, use the required hash-verified retained artifact snapshot. Explicit path tests a new candidate. Declared suite inputs are hash-verified snapshots with fresh working copies; undeclared external files are not frozen. fresh=true uses an independent hidden host. Requires --execute. Each run keeps a distinct saved_report and XML evidence."},
	{"test.compare", "reference! path! xml_policy", "Compare baseline reference and candidate path native report files from the same suite. Automatically compares hash-verified XML snapshots using optional xml_policy, alongside asserted-field parity differences, per-step timing ratios, and separate task_timings for completed asynchronous polls. Poll-call latency is not macro duration. Does not execute Word or certify untested features."},
	{"restore", "path!", "Restore the previous template from an installed activation plan path; refuses newer user edits and waits for Word to be closed."},
	{"compile", "output signing", "Build, compile in real Word and automatically sign a DOTM. Creates a local non-exportable certificate once and reuses it; no certificate parameters needed. Requires --execute. Optional output; cached exact unchanged outputs return immediately."},
	{"sign", "path! output! signing", "Sign an unsigned path to a NEW output DOTM using an automatically created or reused local certificate; optional advanced signing.thumbprint, signing.store_location, signing.signtool and signing.timestamp_url. Requires --execute, Windows SDK SignTool and registered Microsoft Office SIP. Creates legacy, agile and V3 signatures and verifies the strongest V3 digest; publisher trust is reported separately. Run native acceptance on the signed output before deployment."},
	{"journal.catalog", "path! years", "Scan a final-contracts directory for requested current years and return one compact profile per journal: volume/issue evidence, observed fonts/sizes, and feature signals. Reads bounded provenance/summary records only; does not copy article text."},
	{"journal.profile", "journal!", "Return the editable default profile for one supported Canadian law journal. The profile controls generated style names, fonts, sizes, ordered tools and permalink policy."},
	{"journal.create", "journal! output! path years", "Create and offline-build one journal workspace and DOTM from its profile. Optional path names a final-contracts catalog whose observed values are overlaid. The output contains ordinary VBA, a clickable setup form, Ribbon callbacks and no WordUp runtime dependency; native compile/sign/acceptance remain explicit."},
	{"journal.create-all", "output! path years", "Create and offline-build every bundled Canadian law-journal profile beneath output. Optional path supplies one final-contracts catalog scan for all profiles. Existing destinations are never overwritten."},
	{"signature.verify", "path! signing", "Verify the newest VBA signature and certificate trust with Microsoft Office SIP: path, optional signing.signtool. Does not execute macros."},
}

func Tools() []Tool {
	props := map[string]any{}
	props["cases"] = map[string]any{"type": "array", "minItems": 1, "maxItems": 10000, "items": map[string]any{"type": "object", "properties": map[string]any{"id": map[string]any{"type": "string"}, "inputs": map[string]any{"type": "object", "additionalProperties": map[string]any{"type": "string"}}}, "required": []string{"id", "inputs"}, "additionalProperties": false}}
	props["recipe"] = map[string]any{"type": "object", "properties": map[string]any{"styles": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}, "numbering": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}}, "additionalProperties": false, "description": "styles: [{id,name,type,based_on,next,linked,quick,run,paragraph,xml}]; numbering: [{id,levels:[{level,start,format,text,suffix,run,paragraph}]}]. Select only the recipe matching the source part. Unknown properties are rejected; see docs/API.md for units and supported formatting keys."}
	props["comparison"] = map[string]any{"type": "string", "enum": []string{"exact", "semantic"}, "description": "xml.verify comparison; defaults to exact bytes."}
	props["patches"] = map[string]any{"type": "array", "minItems": 1, "maxItems": 65536, "items": map[string]any{"type": "object", "properties": map[string]any{"offset": map[string]any{"type": "integer", "minimum": 0}, "length": map[string]any{"type": "integer", "minimum": 0}, "text": map[string]any{"type": "string"}}, "required": []string{"offset", "length", "text"}, "additionalProperties": false}}
	props["namespaces"] = map[string]any{"type": "object", "additionalProperties": map[string]any{"type": "string"}, "description": "XPath prefix to namespace URI bindings, independent of source prefixes."}
	props["xml_policy"] = map[string]any{"type": "object", "description": "Explicit ignored XML attributes/elements, each a {Space: namespace URI, Local: local name} object. No automatic exclusions.", "properties": map[string]any{"attributes": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}, "elements": map[string]any{"type": "array", "items": map[string]any{"type": "object"}}}, "additionalProperties": false}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["context_parts"] = map[string]any{"type": "array", "items": map[string]any{"type": "string"}, "description": "Flat OPC part names to prioritize for bounded difference context (up to eight); this does not filter comparison or change equality."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["generated_toc_bookmarks"] = map[string]any{"type": "boolean", "description": "Opt-in canonical names for generated _Toc<number> bookmarks, hyperlink anchors and field references, by bookmark document order. Preserves positions, content and link correspondence; raw bytes/hashes remain unchanged."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["generated_comment_ids"] = map[string]any{"type": "boolean", "description": "Opt-in mapping of modern comment paragraph and durable IDs through comment identity. Preserves reply parents, resolution state, text and timestamps; rejects unresolved links. Supply the complete Flat OPC XML including comment parts."}
	props["xml_policy"].(map[string]any)["properties"].(map[string]any)["attribute_equivalences"] = map[string]any{"type": "array", "description": "Explicit equivalent values for one XML element/attribute pair. Each rule has element and attribute {Space,Local}, and values:string[]. Only listed values are equivalent; other elements and values stay compared.", "items": map[string]any{"type": "object", "properties": map[string]any{"element": map[string]any{"type": "object"}, "attribute": map[string]any{"type": "object"}, "values": map[string]any{"type": "array", "items": map[string]any{"type": "string"}}}, "required": []string{"element", "attribute", "values"}, "additionalProperties": false}}
	props["compilation_constants"] = map[string]any{"type": "object", "description": "VBA conditional compilation constants; values are Boolean, string, or number. Defaults target the local 64-bit platform and VBA7."}
	props["paths"] = map[string]any{"type": "array", "items": map[string]any{"type": "string"}}
	props["signing"] = map[string]any{"type": "object", "properties": map[string]any{"thumbprint": map[string]any{"type": "string"}, "signtool": map[string]any{"type": "string"}, "timestamp_url": map[string]any{"type": "string"}, "store_location": map[string]any{"type": "string"}}, "additionalProperties": false}
	props["native_options"] = map[string]any{"type": "object", "description": "native.start options: memory_limit_mb (default 2048), cpu_percent (default 50), startup_timeout_ms. Effective at host creation only."}
	for _, key := range []string{"path", "output", "name", "text", "base64", "expected_sha256", "query", "reference", "document", "part"} {
		props[key] = map[string]any{"type": "string"}
	}
	props["component"] = map[string]any{"type": "string", "description": "Stable bundled component ID such as structure.detect."}
	props["journal"] = map[string]any{"type": "string", "description": "Stable journal profile ID such as ALTA-L-REV or UBC-L-REV."}
	props["years"] = map[string]any{"type": "array", "items": map[string]any{"type": "integer"}, "description": "Current years for journal.catalog; defaults to 2025 and 2026."}
	props["parameters"] = map[string]any{"type": "object", "additionalProperties": map[string]any{"type": "string"}, "description": "Declared component adaptation values. Currently module_prefix is a validated VBA identifier; undeclared values are rejected."}
	for _, key := range []string{"timeout_ms", "tolerance", "offset", "limit", "count"} {
		props[key] = map[string]any{"type": "integer"}
	}
	props["fresh"] = map[string]any{"type": "boolean"}
	props["paused"] = map[string]any{"type": "boolean", "description": "For vba.immediate, evaluate against the currently paused VBA frame through the owned VBE Immediate Window instead of a scratch project."}
	props["symbols"] = map[string]any{"type": "array", "items": map[string]any{"type": "string"}, "description": "Optional paused-frame names to inspect individually; unavailable names remain explicit rows."}
	props["operation"] = map[string]any{"type": "object", "description": NativeHelp}
	props["suite"] = map[string]any{"type": "object", "description": "schema=1,name,steps:[{name,operation,assert:[{kind,path,expected,tolerance}],timeout_ms}], optional platforms, require_compile, and serial_trace (instrumented trace .json filename below $output; validated, retained and compared in order)"}
	tools := make([]Tool, 0, len(toolSpecs))
	for _, spec := range toolSpecs {
		if hiddenByProfile(spec.name) {
			continue
		}
		selected := map[string]any{}
		required := []string{}
		for _, field := range strings.Fields(spec.parameters) {
			name := strings.TrimSuffix(field, "!")
			selected[name] = props[name]
			if name != field {
				required = append(required, name)
			}
		}
		schema := map[string]any{"type": "object", "properties": selected, "additionalProperties": false}
		if len(required) > 0 {
			schema["required"] = required
		}
		tools = append(tools, Tool{spec.name, spec.description, schema})
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
