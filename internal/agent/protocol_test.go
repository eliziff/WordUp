package agent

import (
	"bytes"
	"context"
	"encoding/json"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"strings"
	"testing"
)

func TestMCPHandshakeAndExplicitErrors(t *testing.T) {
	input := `{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-11-25","clientInfo":{"name":"test","version":"1"},"capabilities":{}}}
{"jsonrpc":"2.0","method":"notifications/initialized"}
{"jsonrpc":"2.0","id":2,"method":"tools/list"}
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"read","arguments":{"path":"../escape"}}}
{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"selftest","arguments":{}}}
`
	var output bytes.Buffer
	e := &Engine{Root: t.TempDir()}
	if err := Serve(context.Background(), e, strings.NewReader(input), &output, true); err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimSpace(output.String()), "\n")
	if len(lines) != 4 {
		t.Fatalf("unexpected protocol output %s", output.String())
	}
	var r map[string]any
	for i, line := range lines {
		if err := json.Unmarshal([]byte(line), &r); err != nil {
			t.Fatal(err)
		}
		if r["jsonrpc"] != "2.0" {
			t.Fatal("bad JSONRPC")
		}
		if i >= 2 {
			v := r["result"].(map[string]any)
			if v["isError"] != true {
				t.Fatal("unsafe operation not reported as tool error")
			}
		}
	}
}
func TestProtocolWriteReadHashGuard(t *testing.T) {
	root := t.TempDir() + "/work"
	if _, e := project.New("Test", root); e != nil {
		t.Fatal(e)
	}
	e := &Engine{Root: root}
	defer e.Close()
	ctx := context.Background()
	if _, err := e.Call(ctx, "write", Parameters{Path: "vba/Helpers.bas", Text: "Option Explicit\n"}); err != nil {
		t.Fatal(err)
	}
	v, err := e.Call(ctx, "read", Parameters{Path: "vba/Helpers.bas"})
	if err != nil {
		t.Fatal(err)
	}
	b, _ := json.Marshal(v)
	if !bytes.Contains(b, []byte("sha256")) {
		t.Fatal("missing source hash")
	}
	if _, err = e.Call(ctx, "write", Parameters{Path: "vba/Helpers.bas", Text: "bad", ExpectedSHA256: "wrong"}); err == nil {
		t.Fatal("stale agent write accepted")
	}
}

func TestProtocolReadRangeRetainsFullFileHash(t *testing.T) {
	root := t.TempDir() + "/work"
	if _, e := project.New("Test", root); e != nil {
		t.Fatal(e)
	}
	if err := project.Write(root, "reports/evidence.xml", []byte("0123456789"), ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	value, err := e.Call(context.Background(), "read", Parameters{Path: "reports/evidence.xml", Offset: 3, Limit: 4})
	if err != nil {
		t.Fatal(err)
	}
	row := value.(map[string]any)
	if row["text"] != "3456" || row["bytes"] != 10 || row["offset"] != 3 || row["length"] != 4 || row["truncated"] != true || row["next_offset"] != 7 {
		t.Fatalf("invalid bounded read: %#v", row)
	}
	if row["sha256"] != office.Hash([]byte("0123456789")) {
		t.Fatal("bounded read lost full-file identity")
	}
}
