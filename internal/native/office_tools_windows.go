//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"time"
)

func OfficeTools(ctx context.Context, args ...string) (any, error) {
	exe, err := os.Executable()
	if err != nil {
		return nil, err
	}
	helper := filepath.Join(filepath.Dir(exe), "office-tools", "WordUp.OfficeTools.exe")
	if len(args) > 0 && args[0] == "ole.inspect" {
		helper = filepath.Join(filepath.Dir(exe), "wordup-oletools", "wordup-oletools.exe")
		args = args[1:]
	}
	if _, err = os.Stat(helper); err != nil {
		return nil, Fail("office_tools_missing", "Bundled Office tools helper is missing; build tools/office-bridge/build.ps1 or use the complete distribution", nil)
	}
	ctx, cancel := context.WithTimeout(ctx, 20*time.Second)
	defer cancel()
	result, err := callOfficeWorker(ctx, helper, args)
	if err != nil {
		return nil, err
	}
	if result["error"] != nil {
		return nil, Fail("office_tools_failed", fmt.Sprint(result["error"]), result)
	}
	return result, nil
}
