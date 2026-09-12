// Package deploy stages only artifact-bound native acceptance and never closes
// the user's Word. Queued activation waits locally; it is not a system service.
package deploy

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"
	"wordwright.local/internal/office"
	"wordwright.local/internal/project"
	"wordwright.local/internal/verify"
)

type Plan struct {
	Schema       int    `json:"schema"`
	State        string `json:"state"`
	Candidate    string `json:"candidate"`
	SHA256       string `json:"sha256"`
	Target       string `json:"target"`
	PriorSHA256  string `json:"prior_sha256"`
	Backup       string `json:"backup,omitempty"`
	ReportSHA256 string `json:"report_sha256"`
	Created      string `json:"created"`
	Installed    string `json:"installed,omitempty"`
	Error        string `json:"error,omitempty"`
}

func Prepare(root, artifact, report, target string) (string, *Plan, error) {
	source, e := project.Read(filepath.Dir(artifact), filepath.Base(artifact))
	if e != nil {
		return "", nil, e
	}
	proof, e := project.Read(filepath.Dir(report), filepath.Base(report))
	if e != nil {
		return "", nil, e
	}
	var r verify.Report
	if e = project.ReadJSON(proof, &r); e != nil {
		return "", nil, e
	}
	if r.Status != "passed" || !r.WordExecuted || !r.FreshProcess || r.SHA256 != office.Hash(source) || r.Assertions < 1 || r.SuiteSHA256 != office.Hash(project.JSON(r.Suite)) {
		return "", nil, fmt.Errorf("activation requires fresh, passing native evidence bound to these exact artifact bytes and suite")
	}
	if e = r.Suite.Validate(); e != nil {
		return "", nil, e
	}
	if len(r.Observations) != len(r.Suite.Steps) {
		return "", nil, fmt.Errorf("incomplete observation set")
	}
	n := 0
	for i, o := range r.Observations {
		if !o.Passed || o.Name != r.Suite.Steps[i].Name || o.Assertions != len(r.Suite.Steps[i].Assert) {
			return "", nil, fmt.Errorf("native observation mismatch")
		}
		n += o.Assertions
	}
	if n != r.Assertions {
		return "", nil, fmt.Errorf("assertion count mismatch")
	}
	target, e = filepath.Abs(target)
	if e != nil {
		return "", nil, e
	}
	if !strings.EqualFold(filepath.Ext(target), ".dotm") {
		return "", nil, fmt.Errorf("deployment target must be .dotm")
	}
	prior := "new-file"
	if b, e := project.Read(filepath.Dir(target), filepath.Base(target)); e == nil {
		prior = office.Hash(b)
	} else if !os.IsNotExist(e) {
		return "", nil, e
	}
	base := filepath.Join(root, ".wordwright", "deploy")
	if e = os.MkdirAll(base, 0700); e != nil {
		return "", nil, e
	}
	dir, e := os.MkdirTemp(base, "release-")
	if e != nil {
		return "", nil, e
	}
	candidate := filepath.Join(dir, "candidate.dotm")
	if e = project.AtomicWrite(candidate, source); e != nil {
		return "", nil, e
	}
	if e = project.AtomicWrite(filepath.Join(dir, "acceptance.json"), proof); e != nil {
		return "", nil, e
	}
	p := &Plan{Schema: 1, State: "pending", Candidate: candidate, SHA256: office.Hash(source), Target: target, PriorSHA256: prior, ReportSHA256: office.Hash(proof), Created: time.Now().UTC().Format(time.RFC3339Nano)}
	file := filepath.Join(dir, "activation.json")
	if e = project.AtomicWrite(file, project.JSON(p)); e != nil {
		return "", nil, e
	}
	return file, p, nil
}
func Activate(file string) (*Plan, error) {
	b, e := project.Read(filepath.Dir(file), filepath.Base(file))
	if e != nil {
		return nil, e
	}
	p := &Plan{}
	if e = project.ReadJSON(b, p); e != nil {
		return nil, e
	}
	if p.Schema != 1 {
		return p, fmt.Errorf("activation schema mismatch")
	}
	if p.State != "pending" {
		return p, fmt.Errorf("activation is not pending")
	}
	busy, e := WordRunning()
	if e != nil {
		return p, e
	}
	if busy {
		return p, fmt.Errorf("word_still_running")
	}
	source, e := project.Read(filepath.Dir(p.Candidate), filepath.Base(p.Candidate))
	if e != nil {
		return p, e
	}
	if office.Hash(source) != p.SHA256 {
		return p, fmt.Errorf("staged candidate changed")
	}
	proof, e := project.Read(filepath.Dir(file), "acceptance.json")
	if e != nil || office.Hash(proof) != p.ReportSHA256 {
		return p, fmt.Errorf("staged acceptance record changed")
	}
	prior, e := project.Read(filepath.Dir(p.Target), filepath.Base(p.Target))
	if p.PriorSHA256 == "new-file" {
		if !os.IsNotExist(e) {
			return p, fmt.Errorf("target now exists; refusing to overwrite a concurrent installation")
		}
	} else {
		if e != nil || office.Hash(prior) != p.PriorSHA256 {
			return p, fmt.Errorf("installed template changed after staging; refusing to overwrite newer user changes")
		}
		backup := filepath.Join(filepath.Dir(file), "previous.dotm")
		if e = project.AtomicWrite(backup, prior); e != nil {
			return p, e
		}
		p.Backup = backup
	}
	if e = project.AtomicWrite(p.Target, source); e != nil {
		return p, e
	}
	p.State = "installed"
	p.Installed = time.Now().UTC().Format(time.RFC3339Nano)
	if e = project.AtomicWrite(file, project.JSON(p)); e != nil {
		return p, e
	}
	return p, nil
}
func Watch(ctx context.Context, file string) error {
	ticker := time.NewTicker(500 * time.Millisecond)
	defer ticker.Stop()
	for {
		busy, e := WordRunning()
		if e != nil {
			return e
		}
		if !busy {
			_, e = Activate(file)
			return e
		}
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-ticker.C:
		}
	}
}
