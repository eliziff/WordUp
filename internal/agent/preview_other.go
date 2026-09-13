//go:build !windows

package agent

import (
	"context"
	"fmt"
)

func preview(context.Context, string, string, string, string) (any, error) {
	return nil, fmt.Errorf("visible preview is currently implemented on Windows only")
}
