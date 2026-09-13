//go:build !windows || (!amd64 && !arm64)

package native

import (
	"context"
	"fmt"
)

func OfficeTools(context.Context, ...string) (any, error) {
	return nil, fmt.Errorf("Office tools bridge currently supports 64-bit Windows")
}

func StopOfficeTools() {}
