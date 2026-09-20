package agent

import (
	"context"
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
	hash := office.Hash(b)
	evidence := map[string]any{"status": "not_provided"}
	if proof != "" {
		p, err := project.Read(filepath.Dir(proof), filepath.Base(proof))
		if err != nil {
			return nil, err
		}
		var report verify.Report
		if err := verify.DecodeReport(p, &report); err != nil {
			return nil, err
		}
		evidence = map[string]any{"reference": proof, "reported_status": report.Status,
			"artifact_matches": report.SHA256 == hash, "reported_vba_compiled": report.VBACompiled,
			"reported_fresh_process": report.FreshProcess}
	}
	sourceDocument := document
	// Word can return VT_EMPTY rather than an error for an overlong path.
	// Runtime copies must not inherit arbitrary repository depth or input names.
	cache, err := os.UserCacheDir()
	if err != nil {
		return nil, err
	}
	parent := filepath.Join(cache, "WordUp", "previews")
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
	if document != "" {
		manuscript, readErr := project.Read(filepath.Dir(document), filepath.Base(document))
		if readErr != nil {
			return nil, readErr
		}
		document = filepath.Join(directory, "manuscript"+filepath.Ext(document))
		if err = project.AtomicWrite(document, manuscript); err != nil {
			return nil, err
		}
	}
	result, err := native.Preview(ctx, destination, document)
	if err != nil {
		return result, err
	}
	result.(map[string]any)["acceptance_reference"] = evidence
	result.(map[string]any)["sha256"] = hash
	result.(map[string]any)["source_document"] = sourceDocument
	return result, nil
}
