//go:build !windows

package signing

import (
	"context"
	"fmt"
	"os/exec"
)

func hide(c *exec.Cmd)                     {}
func VerifyDigestWorker(file string) error { return fmt.Errorf("Office SIP requires Windows") }

func EnsureCertificate(ctx context.Context) (*Certificate, error) {
	return nil, fmt.Errorf("automatic VBA certificates require Windows")
}
