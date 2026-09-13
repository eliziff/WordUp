package agent

import (
	"context"
	"fmt"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"os"
	"path/filepath"
)

func preview(ctx context.Context, root, artifact, proof, document string) (any, error) {
	b, err := project.Read(filepath.Dir(artifact), filepath.Base(artifact))
	if err != nil {
		return nil, err
	}
	p, err := project.Read(filepath.Dir(proof), filepath.Base(proof))
	if err != nil {
		return nil, err
	}
	var report verify.Report
	if err = verify.DecodeReport(p, &report); err != nil {
		return nil, err
	}
	if _, err = verify.Compare(&report, &report); err != nil {
		return nil, err
	}
	hash := office.Hash(b)
	if report.SHA256 != hash || !report.FreshProcess || !report.VBACompiled {
		return nil, fmt.Errorf("preview requires fresh native compilation and acceptance for these exact bytes")
	}
	parent := filepath.Join(root, ".wordup", "preview", hash)
	if err = os.MkdirAll(parent, 0700); err != nil {
		return nil, err
	}
	directory, err := os.MkdirTemp(parent, "run-")
	if err != nil {
		return nil, err
	}
	destination := filepath.Join(directory, "preview.dotm")
	if err = project.AtomicWrite(destination, b); err != nil {
		return nil, err
	}
	return native.Preview(ctx, destination, document)
}
