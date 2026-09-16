package agent

import (
	"fmt"
	"strings"
	"sync"
)

// Tool profiles let a benchmark or an operator expose only the core
// infrastructure surface. "core" hides the model-facing abstractions
// (bundled VBA components, journal generation and structure detection) so an
// agent has to author the equivalent VBA itself; "" or "full" exposes
// everything. The profile applies to the manifest and to dispatch.
var (
	profileMu   sync.RWMutex
	toolProfile string
)

func SetToolProfile(profile string) error {
	switch profile {
	case "", "full":
		profile = ""
	case "core":
	default:
		return fmt.Errorf("unknown tool profile %q; use core or full", profile)
	}
	profileMu.Lock()
	toolProfile = profile
	profileMu.Unlock()
	return nil
}

func ToolProfile() string {
	profileMu.RLock()
	defer profileMu.RUnlock()
	if toolProfile == "" {
		return "full"
	}
	return toolProfile
}

func hiddenByProfile(name string) bool {
	profileMu.RLock()
	profile := toolProfile
	profileMu.RUnlock()
	if profile != "core" {
		return false
	}
	for _, prefix := range []string{"component.", "journal.", "structure."} {
		if strings.HasPrefix(name, prefix) {
			return true
		}
	}
	return false
}
