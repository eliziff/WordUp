package agent

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/signing"
	"github.com/eliziff/WordUp/internal/verify"
)

type CompileReport struct {
	Build      *project.BuildReport `json:"build"`
	Signing    *signing.Report      `json:"signing"`
	DurationMS float64              `json:"duration_ms"`
	Cached     bool                 `json:"cached"`
	Native     *verify.Report       `json:"native,omitempty"`
	Error      any                  `json:"error,omitempty"`
}

func (e *Engine) compile(ctx context.Context, w *project.Workspace, output string, opts signing.Options) (report *CompileReport, retErr error) {
	if !e.Execute {
		return nil, fmt.Errorf("compile requires --execute to run Word and sign the output")
	}
	start := time.Now()
	if output == "" {
		output = filepath.Join(w.Root, "dist", w.Manifest.Name+".dotm")
	}
	output, err := filepath.Abs(output)
	if err != nil {
		return nil, err
	}
	files, err := w.SourceFiles()
	if err != nil {
		return nil, err
	}
	fingerprint := project.Fingerprint(files)
	if b, err := project.Read(w.Root, "reports/compile.json"); err == nil && opts == (signing.Options{}) {
		var cached CompileReport
		// A cached signature must still have a valid app-created certificate.
		cacheErr := project.ReadJSON(b, &cached)
		var expires time.Time
		if cached.Signing != nil && cached.Signing.Certificate != nil {
			expires, _ = time.Parse(time.RFC3339Nano, cached.Signing.Certificate.Expires)
		}
		if cacheErr == nil && time.Now().Before(expires) && cached.Error == nil && cached.Build != nil && cached.Build.ToolVersion == project.Version && cached.Signing != nil && cached.Signing.Signed && cached.Signing.DigestVerified && cached.Build.VBACompiled && cached.Build.Artifact == output && cached.Build.SourceFingerprint == fingerprint {
			if artifact, err := os.ReadFile(output); err == nil && office.Hash(artifact) == cached.Build.SHA256 && cached.Signing.SHA256 == cached.Build.SHA256 {
				cached.Cached = true
				cached.DurationMS = float64(time.Since(start).Microseconds()) / 1000
				return &cached, nil
			}
		}
	}
	if err = os.MkdirAll(filepath.Dir(output), 0700); err != nil {
		return nil, err
	}
	prior, priorErr := os.ReadFile(output)
	if priorErr != nil && !os.IsNotExist(priorErr) {
		return nil, priorErr
	}
	stage, err := os.MkdirTemp(filepath.Dir(output), ".wordup-compile-")
	if err != nil {
		return nil, err
	}
	// Keep failed candidates and visual diagnostics available to the agent.
	defer func() {
		if retErr == nil {
			_ = os.RemoveAll(stage)
		}
	}()
	build, err := w.Build(filepath.Join(stage, "unsigned.dotm"))
	if err != nil {
		return nil, err
	}
	r := &CompileReport{Build: build}
	defer func() {
		r.DurationMS = float64(time.Since(start).Microseconds()) / 1000
		r.Error = verify.ErrorValue(retErr)
		if saveErr := project.Write(w.Root, "reports/compile.json", project.JSON(r), ""); retErr == nil {
			retErr = saveErr
		}
	}()
	h, err := e.Host(ctx)
	if err != nil {
		return r, err
	}
	ctx, cancel := context.WithTimeout(ctx, 60*time.Second)
	defer cancel()
	r.Native, err = verify.Run(ctx, build.Artifact, verify.Suite{Schema: 1, Name: "Compile candidate before signing", RequireCompile: true, Steps: []verify.Step{
		{Name: "Open candidate", Operation: native.Operation{Op: "open", File: "$artifact", As: "compileCandidate"}},
		{Name: "Compile entire project", Operation: native.Operation{Op: "compile", Target: "compileCandidate", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Return to Word", Operation: native.Operation{Op: "put", Target: "app", Member: "ShowVisualBasicEditor", Value: false}},
		{Name: "Unload candidate", Operation: native.Operation{Op: "unload", Target: "compileCandidate", File: "$artifact"}},
	}}, h, true)
	if err != nil {
		_ = e.Close()
		return r, err
	}
	build.VBACompiled = true
	build.WordExecuted = true
	r.Signing, err = signing.Sign(ctx, build.Artifact, filepath.Join(stage, "signed.dotm"), opts)
	if err != nil {
		return r, err
	}
	signed, err := os.ReadFile(r.Signing.Artifact)
	if err != nil {
		return r, err
	}
	now, nowErr := os.ReadFile(output)
	if (priorErr == nil && (nowErr != nil || office.Hash(now) != office.Hash(prior))) || (os.IsNotExist(priorErr) && !os.IsNotExist(nowErr)) {
		return r, fmt.Errorf("output changed during compilation; newer file preserved")
	}
	if err = project.AtomicWrite(output, signed); err != nil {
		return r, err
	}
	build.Artifact = output
	build.SHA256 = office.Hash(signed)
	build.Bytes = len(signed)
	build.NoChange = false
	r.Signing.Artifact = output
	r.DurationMS = float64(time.Since(start).Microseconds()) / 1000
	return r, nil
}
