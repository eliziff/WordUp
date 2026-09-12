package agent

import (
	"bufio"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"strings"
	"time"
	"wordwright.local/internal/project"
	"wordwright.local/internal/verify"
)

type Tool struct {
	Name        string         `json:"name"`
	Description string         `json:"description"`
	InputSchema map[string]any `json:"inputSchema"`
}

func Tools() []Tool {
	ds := [][2]string{{"example", "Create and build an integrated editable native template: output."}, {"selftest", "Generate own integration fixture and test it in actual local Word; saves assertions/failures. Requires --execute, optional output."}, {"deploy", "Stage exact native-tested artifact path, acceptance reference, destination output. Automatically waits locally for Word to close; backup and stale-target guards."}, {"activate", "Retry a durable activation plan path after an interrupted local worker. Never closes Word."}, {"help", "Read source layout, operations and native ABI help."}, {"doctor", "Detect local native Word. Does not execute or emulate it."}, {"new", "Create a native source workspace: name,output."}, {"import", "Import an actual DOCX/DOTM package: path,output."}, {"inspect", "Inventory an artifact: path."}, {"files", "List source files; limit bounds the response."}, {"read", "Read workspace path and its optimistic concurrency hash."}, {"write", "Write UTF-8 text or base64 to path; expected_sha256 protects existing edits."}, {"search", "Literal case-insensitive source search: query,limit."}, {"build", "Build a native DOTM; optional output. No VBA execution claim."}, {"check", "Fast lexical/XML diagnostics, NOT full compilation."}, {"compat", "Conservative conditional Windows/Mac portability review, NOT a Mac pass."}, {"reference.document", "Read native styles and paragraph/property observations from path."}, {"reference.image", "Return reference/screenshot image pixels to the vision-capable agent."}, {"image.compare", "Exact-size pixel comparison: reference,path,tolerance,optional output difference image."}, {"native.start", "Start owned local Word lazily; requires installed Word, not a VM."}, {"native.call", "Raw native operation; see help. Persistent host. Macro execution requires launch --execute."}, {"native.stop", "Close ONLY the app-owned Word and private desktop."}, {"test", "Execute assertion suite against exact artifact path; fresh=true uses an independent process. Returns and saves truthful pass/fail/not_run evidence."}}
	props := map[string]any{}
	for _, key := range []string{"path", "output", "name", "text", "base64", "expected_sha256", "query", "reference"} {
		props[key] = map[string]any{"type": "string"}
	}
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
				if err := emit(q.ID, map[string]any{"protocolVersion": version, "serverInfo": map[string]any{"name": "wordwright", "version": project.Version}, "capabilities": map[string]any{"tools": map[string]any{}}, "instructions": project.AgentInstructions}, nil, 0); err != nil {
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
