//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"github.com/eliziff/WordUp/internal/native"
)

// Headless Word launch is a core product requirement. This soak launches
// owned hidden Word repeatedly, then repeats a smaller batch while another
// owned Word is already running, and fails on any launch or cleanup fault.
// Set WORDUP_LAUNCH_SOAK to change the cold count (default 20).
func TestNativeLaunchSoak(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	count := 20
	if v, err := strconv.Atoi(os.Getenv("WORDUP_LAUNCH_SOAK")); err == nil && v > 0 {
		count = v
	}
	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Minute)
	defer cancel()
	dir := t.TempDir()
	cold, err := native.Probe(ctx, native.Options{Directory: filepath.Join(dir, "cold")}, count)
	report(t, "cold", cold)
	if err != nil {
		t.Fatalf("cold launch soak failed: %v", err)
	}
	if cold["desktop_isolated"] != count {
		t.Fatalf("desktop isolation held for %v of %d cold launches", cold["desktop_isolated"], count)
	}
	// A second owned session while the first stays open models a user's Word
	// already running: the new launches must not attach to it or be refused.
	resident, err := native.Start(ctx, native.Options{Directory: filepath.Join(dir, "resident")})
	if err != nil {
		t.Fatalf("resident session: %v", err)
	}
	defer resident.Close()
	concurrent, err := native.Probe(ctx, native.Options{Directory: filepath.Join(dir, "concurrent")}, 5)
	report(t, "concurrent", concurrent)
	if err != nil {
		t.Fatalf("launch soak with another Word open failed: %v", err)
	}
	if info := resident.Info(); info["closed"] != false {
		t.Fatalf("resident session was disturbed: %v", info)
	}
}

func report(t *testing.T, name string, m map[string]any) {
	t.Helper()
	if m == nil {
		return
	}
	summary := map[string]any{}
	for _, k := range []string{"requested", "launched", "failed", "close_errors", "desktop_isolated", "word_versions", "startup_ms_min", "startup_ms_p50", "startup_ms_p95", "startup_ms_max"} {
		summary[k] = m[k]
	}
	b, _ := json.Marshal(summary)
	t.Logf("%s: %s", name, b)
	if out := os.Getenv("WORDUP_LAUNCH_SOAK_REPORT"); out != "" {
		full, _ := json.MarshalIndent(m, "", "  ")
		_ = os.WriteFile(filepath.Join(out, name+"-launch-soak.json"), full, 0644)
	}
}
