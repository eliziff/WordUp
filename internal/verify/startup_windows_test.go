//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"os"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
)

func TestNativeProbeRetainsStartupDiagnostics(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	report, err := native.Probe(context.Background(), native.Options{}, 1)
	if err != nil {
		t.Fatal(err)
	}
	launch := report["launches"].([]map[string]any)[0]
	diagnostics, _ := launch["startup_diagnostics"].(string)
	previous := -1
	for _, milestone := range []string{"spawned WINWORD", "document window found", "native accessibility returned error=<nil>", "ownership checked", "Application acquired", "application settings complete", "hello: ready"} {
		position := strings.Index(diagnostics, milestone)
		if position <= previous {
			t.Fatalf("missing or unordered milestone %q: %s", milestone, diagnostics)
		}
		previous = position
	}
}
