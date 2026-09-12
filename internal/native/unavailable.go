//go:build !windows && !darwin

package native

import (
	"context"
	"fmt"
	"runtime"
)

func Available() bool { return false }
func Start(context.Context, Options) (Host, error) {
	return nil, Fail("native_runtime_unavailable", fmt.Sprintf("%s has no Microsoft Word runtime; no test was executed", runtime.GOOS), nil)
}
func HostMain([]string) error {
	return Fail("native_runtime_unavailable", "Windows/macOS host required", nil)
}
