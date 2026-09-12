// WordUp is one executable. Native host subprocesses are this same image.
package main

import (
	"context"
	"encoding/json"
	"fmt"
	"github.com/eliziff/WordUp/internal/agent"
	"github.com/eliziff/WordUp/internal/deploy"
	"github.com/eliziff/WordUp/internal/localipc"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/signing"
	"github.com/eliziff/WordUp/internal/verify"
	"io"
	"os"
	"os/signal"
	"path/filepath"
	"strings"
	"time"
)

func main() {
	if len(os.Args) == 3 && os.Args[1] == "__verify_vba_digest" {
		if err := signing.VerifyDigestWorker(os.Args[2]); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		return
	}
	if len(os.Args) > 1 && os.Args[1] == "__host" {
		if e := native.HostMain(os.Args[2:]); e != nil {
			fmt.Fprintln(os.Stderr, e)
			os.Exit(1)
		}
		return
	}
	if len(os.Args) == 3 && os.Args[1] == "__activate" {
		if err := deploy.Watch(context.Background(), os.Args[2]); err != nil {
			deploy.RecordFailure(os.Args[2], err)
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		return
	}
	if len(os.Args) > 1 && os.Args[1] == "__session" {
		if len(os.Args) < 4 || len(os.Args) > 5 {
			fmt.Fprintln(os.Stderr, "invalid local session invocation")
			os.Exit(1)
		}
		execute := len(os.Args) == 5 && os.Args[4] == "execute"
		if err := localipc.Serve(context.Background(), os.Args[2], os.Args[3], execute); err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		return
	}
	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt)
	defer cancel()
	code := run(ctx, os.Args[1:])
	if code != 0 {
		os.Exit(code)
	}
}
func run(ctx context.Context, args []string) int {
	root := "."
	execute := false
	clean := []string{}
	for i := 0; i < len(args); i++ {
		switch args[i] {
		case "--workspace", "-w":
			i++
			if i == len(args) {
				fmt.Fprintln(os.Stderr, "--workspace needs a path")
				return 2
			}
			root = args[i]
		case "--execute":
			execute = true
		default:
			clean = append(clean, args[i])
		}
	}
	if len(clean) == 0 || clean[0] == "--help" || clean[0] == "help" {
		fmt.Print(usage)
		return 0
	}
	if clean[0] == "version" || clean[0] == "--version" {
		fmt.Println(project.Version)
		return 0
	}
	var err error
	root, err = filepath.Abs(root)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		return 1
	}
	e := &agent.Engine{Root: root, Execute: execute}
	defer e.Close()
	if clean[0] == "serve" || clean[0] == "mcp" {
		if err = agent.Serve(ctx, e, os.Stdin, os.Stdout, clean[0] == "mcp"); err != nil {
			fmt.Fprintln(os.Stderr, err)
			return 1
		}
		return 0
	}
	if clean[0] == "session" {
		var result any
		var err error
		if len(clean) != 2 {
			err = fmt.Errorf("session start|info|stop required")
		} else if clean[1] == "start" {
			result, err = localipc.Start(ctx, root, execute)
		} else if clean[1] == "stop" || clean[1] == "info" {
			c, cancel := context.WithTimeout(ctx, 10*time.Second)
			defer cancel()
			result, err = localipc.Call(c, root, "session."+clean[1], agent.Parameters{})
		} else {
			err = fmt.Errorf("unknown session action")
		}
		if err != nil {
			json.NewEncoder(os.Stdout).Encode(map[string]any{"error": verify.ErrorValue(err)})
			return 1
		}
		json.NewEncoder(os.Stdout).Encode(result)
		return 0
	}
	remote := false
	if clean[0] == "rpc" {
		if len(clean) != 3 {
			fmt.Fprintln(os.Stderr, "rpc METHOD JSON|@parameters.json required")
			return 2
		}
		clean[0] = "call"
		remote = true
	}
	method := clean[0]
	p := agent.Parameters{}
	rest := clean[1:]
	getJSON := func(s string) error {
		var b []byte
		var x error
		if s == "-" {
			b, x = io.ReadAll(io.LimitReader(os.Stdin, 16<<20))
		} else if strings.HasPrefix(s, "@") {
			b, x = os.ReadFile(s[1:])
		} else {
			b = []byte(s)
		}
		if x != nil {
			return x
		}
		return project.ReadJSON(b, &p)
	}
	switch method {
	case "call":
		if len(rest) != 2 {
			err = fmt.Errorf("call METHOD JSON or @parameters.json required")
		} else {
			method = rest[0]
			err = getJSON(rest[1])
		}
	case "new":
		if len(rest) != 2 {
			err = fmt.Errorf("new NAME OUTPUT_DIRECTORY required")
		} else {
			p.Name, p.Output = rest[0], rest[1]
		}
	case "import":
		if len(rest) != 2 {
			err = fmt.Errorf("import SOURCE OUTPUT_DIRECTORY required")
		} else {
			p.Path, p.Output = rest[0], rest[1]
		}
	case "inspect", "reference.document", "reference.image", "read":
		if len(rest) != 1 {
			err = fmt.Errorf("%s PATH required", method)
		} else {
			p.Path = rest[0]
		}
	case "example", "selftest":
		if len(rest) > 1 {
			err = fmt.Errorf("%s [OUTPUT_DIRECTORY] expected", method)
		} else if len(rest) == 1 {
			p.Output = rest[0]
		}
		if method == "example" && p.Output == "" {
			err = fmt.Errorf("example OUTPUT_DIRECTORY required")
		}
	case "build", "compile":
		if len(rest) > 1 {
			err = fmt.Errorf("build [OUTPUT] expected")
		} else if len(rest) == 1 {
			p.Output = rest[0]
		}
	case "search":
		if len(rest) != 1 {
			err = fmt.Errorf("search QUERY required")
		} else {
			p.Query = rest[0]
		}
	case "native":
		method = "native.call"
		if len(rest) != 1 {
			err = fmt.Errorf("native JSON_OPERATION or @operation.json required")
		} else {
			var b []byte
			if strings.HasPrefix(rest[0], "@") {
				b, err = os.ReadFile(rest[0][1:])
			} else if rest[0] == "-" {
				b, err = io.ReadAll(io.LimitReader(os.Stdin, 16<<20))
			} else {
				b = []byte(rest[0])
			}
			if err == nil {
				err = project.ReadJSON(b, &p.Operation)
			}
		}
	case "test":
		p.Fresh = true
		if len(rest) < 1 || len(rest) > 2 {
			err = fmt.Errorf("test ARTIFACT [SUITE.json] required")
		} else {
			p.Path = rest[0]
			if len(rest) == 2 {
				b, x := os.ReadFile(rest[1])
				err = x
				if err == nil {
					p.Suite = &verify.Suite{}
					err = project.ReadJSON(b, p.Suite)
				}
			}
		}
	case "doctor", "check", "compat", "files":
		if len(rest) != 0 {
			err = fmt.Errorf("unexpected arguments")
		}
	default:
		err = fmt.Errorf("unknown command %q; use help or call METHOD JSON", method)
	}
	start := time.Now()
	var result any
	if err == nil {
		if remote {
			c, cancel := context.WithTimeout(ctx, 30*time.Minute)
			defer cancel()
			var reply localipc.Reply
			reply, err = localipc.Call(c, root, method, p)
			result = reply.Result
			if err == nil && reply.Error != nil {
				err = fmt.Errorf("local operation failed: %s", project.JSON(reply.Error))
			}
		} else {
			result, err = e.Call(ctx, method, p)
		}
	}
	out := map[string]any{"result": result, "duration_ms": float64(time.Since(start).Microseconds()) / 1000}
	if err != nil {
		out["error"] = verify.ErrorValue(err)
	}
	enc := json.NewEncoder(os.Stdout)
	enc.SetIndent("", "  ")
	if x := enc.Encode(out); x != nil {
		fmt.Fprintln(os.Stderr, x)
		return 1
	}
	if err != nil {
		return 1
	}
	return 0
}

const usage = `WordUp 0.3.0 — native local Word development toolchain

  wordup import SOURCE.dotm WORKSPACE
  wordup new ProjectName WORKSPACE
  wordup --workspace WORKSPACE build [OUTPUT.dotm]
  wordup --workspace WORKSPACE --execute compile [OUTPUT.dotm]
  wordup --workspace WORKSPACE check
  wordup --workspace WORKSPACE compat
  wordup inspect ARTIFACT.dotm
  wordup --workspace WORKSPACE --execute test ARTIFACT.dotm [SUITE.json]
  wordup --workspace WORKSPACE --execute serve
  wordup --workspace WORKSPACE --execute mcp
  wordup --workspace WORKSPACE call METHOD @parameters.json
  wordup --workspace WORKSPACE --execute session start
  wordup --workspace WORKSPACE rpc METHOD @parameters.json
  wordup --workspace WORKSPACE session stop
  wordup doctor
  wordup --execute selftest [EVIDENCE_DIRECTORY]
  wordup example OUTPUT_DIRECTORY

serve: persistent newline JSON {"id":1,"method":"build","params":{}}.
mcp: stdio MCP (2025-03-26 / 2025-06-18 / 2025-11-25 negotiation).
Read AGENTS.md in the imported workspace; call help '{}' lists every tool.
Core authoring needs no Python, Node, cloud service, VM or module imports.
Automatic VBA signing requires Microsoft SignTool and the registered Office SIP.
Native execution uses your installed Microsoft Word and respects OS policies.
Private Windows desktop = UI separation, NOT a security sandbox.
Build/check/compat are NOT VBA runtime tests. --execute grants native code
execution with your account's permissions; authorize only intended code.
`
