package verify

import (
	"context"
	"crypto/sha256"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

// CorpusCase substitutes only declared inputs; steps and assertions stay shared.
type CorpusCase struct {
	ID     string            `json:"id"`
	Inputs map[string]string `json:"inputs"`
}

type CorpusResult struct {
	ID           string `json:"id"`
	Key          string `json:"key"`
	Status       string `json:"status"`
	Report       string `json:"report,omitempty"`
	Bundle       string `json:"bundle,omitempty"`
	BundleSHA256 string `json:"bundle_sha256,omitempty"`
	Reused       bool   `json:"reused,omitempty"`
	Error        string `json:"error,omitempty"`
}

type CorpusReport struct {
	Schema       int            `json:"schema"`
	RunnerSHA256 string         `json:"runner_sha256"`
	Status       string         `json:"status"`
	SavedReport  string         `json:"saved_report"`
	Cases        []CorpusResult `json:"cases"`
}

func corpusRunnerHash() (string, error) {
	path, err := os.Executable()
	if err != nil {
		return "", err
	}
	f, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer f.Close()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return "", err
	}
	return fmt.Sprintf("%x", h.Sum(nil)), nil
}

func corpusIdentity(artifactHash string, suite Suite, inputs map[string]inputBytes) string {
	hashes := map[string]string{}
	for name, input := range inputs {
		hashes[name] = office.Hash(input.data)
	}
	return office.Hash(project.JSON(struct {
		Version  string
		Artifact string
		Suite    Suite
		Inputs   map[string]string
	}{project.Version, artifactHash, suite, hashes}))
}

func (r *CorpusReport) save() error {
	if err := project.AtomicWrite(r.SavedReport, project.JSON(r)); err != nil {
		return err
	}
	var index strings.Builder
	fmt.Fprintf(&index, "# Corpus: %s\n\n", r.Status)
	for _, c := range r.Cases {
		fmt.Fprintf(&index, "- %s: %s", c.ID, c.Status)
		if c.Reused {
			index.WriteString(" (reused verified evidence)")
		}
		if c.Report != "" {
			fmt.Fprintf(&index, " — report: `%s`", c.Report)
		}
		if c.Error != "" {
			fmt.Fprintf(&index, " — %s", strings.ReplaceAll(c.Error, "\n", " "))
		}
		index.WriteByte('\n')
	}
	return project.AtomicWrite(filepath.Join(filepath.Dir(r.SavedReport), "index.md"), []byte(index.String()))
}

// RunCorpus delegates execution to the existing native suite runner. Each call
// allocates a new run directory; retries never overwrite earlier attempts.
// reference is an optional previous corpus summary, not a latest-pointer guess.
func RunCorpus(ctx context.Context, artifact string, suite Suite, cases []CorpusCase, output, reference string, run func(context.Context, string, Suite) (*Report, error)) (*CorpusReport, error) {
	if len(cases) == 0 || len(cases) > 10000 {
		return nil, fmt.Errorf("corpus requires 1..10000 cases")
	}
	if err := suite.Validate(); err != nil {
		return nil, err
	}
	if len(suite.Inputs) == 0 {
		return nil, fmt.Errorf("corpus suite must declare inputs")
	}
	artifactBytes, err := project.Read(filepath.Dir(artifact), filepath.Base(artifact))
	if err != nil {
		return nil, err
	}
	artifactHash := office.Hash(artifactBytes)
	runnerHash, err := corpusRunnerHash()
	if err != nil {
		return nil, err
	}
	suites := make([]Suite, len(cases))
	keys := make([]string, len(cases))
	seen := map[string]bool{}
	for i, c := range cases {
		if !inputName.MatchString(c.ID) || seen[strings.ToLower(c.ID)] {
			return nil, fmt.Errorf("invalid or duplicate corpus case ID %q", c.ID)
		}
		seen[strings.ToLower(c.ID)] = true
		if len(c.Inputs) != len(suite.Inputs) {
			return nil, fmt.Errorf("case %s must replace exactly the declared inputs", c.ID)
		}
		for name := range suite.Inputs {
			if c.Inputs[name] == "" {
				return nil, fmt.Errorf("case %s missing input %s", c.ID, name)
			}
		}
		s := suite
		s.Inputs = c.Inputs
		if err := s.Validate(); err != nil {
			return nil, fmt.Errorf("case %s: %w", c.ID, err)
		}
		inputs, err := readInputs(s.Inputs, nil)
		if err != nil {
			return nil, fmt.Errorf("case %s: %w", c.ID, err)
		}
		suites[i], keys[i] = s, corpusIdentity(artifactHash, s, inputs)
	}
	previous := map[string]CorpusResult{}
	if reference != "" {
		data, err := project.Read(filepath.Dir(reference), filepath.Base(reference))
		if err != nil {
			return nil, err
		}
		var prior CorpusReport
		if err := project.ReadJSON(data, &prior); err != nil {
			return nil, err
		}
		if prior.Schema != 1 {
			return nil, fmt.Errorf("expected schema 1 corpus summary")
		}
		if prior.RunnerSHA256 == runnerHash {
			for _, c := range prior.Cases {
				previous[c.ID] = c
			}
		}
	}
	if err := os.MkdirAll(output, 0700); err != nil {
		return nil, err
	}
	dir, err := os.MkdirTemp(output, "run-")
	if err != nil {
		return nil, err
	}
	r := &CorpusReport{Schema: 1, RunnerSHA256: runnerHash, Status: "incomplete", SavedReport: filepath.Join(dir, "corpus.json")}
	for i, c := range cases {
		r.Cases = append(r.Cases, CorpusResult{ID: c.ID, Key: keys[i], Status: "pending"})
	}
	if err := r.save(); err != nil {
		return r, err
	}
	for i := range cases {
		if err := ctx.Err(); err != nil {
			return r, err
		}
		c := &r.Cases[i]
		if old, ok := previous[c.ID]; ok && old.Key == c.Key && old.Status == "passed" {
			// Reuse only a hash-verified, complete frozen bundle, not a status bit.
			data, readErr := project.Read(filepath.Dir(old.Bundle), filepath.Base(old.Bundle))
			if readErr == nil && office.Hash(data) == old.BundleSHA256 {
				prior, _, loadErr := LoadReport(old.Bundle)
				if loadErr == nil && checkedReport(prior) == nil && corpusReportMatches(prior, artifactHash, suites[i], keys[i]) == nil {
					*c = old
					c.Reused = true
					if err := r.save(); err != nil {
						return r, err
					}
					continue
				}
			}
		}
		c.Status = "running"
		if err := r.save(); err != nil {
			return r, err
		}
		report, runErr := run(ctx, artifact, suites[i])
		if report != nil {
			c.Report = report.SavedReport
		}
		if runErr == nil {
			runErr = checkedReport(report)
		}
		if runErr == nil {
			runErr = corpusReportMatches(report, artifactHash, suites[i], keys[i])
		}
		if runErr == nil {
			bundleDir := filepath.Join(dir, "case-"+c.ID)
			_, runErr = Freeze(report.SavedReport, bundleDir)
			if runErr == nil {
				c.Bundle = filepath.Join(bundleDir, "bundle.json")
				data, readErr := project.Read(bundleDir, "bundle.json")
				runErr = readErr
				c.BundleSHA256 = office.Hash(data)
			}
		}
		c.Status = "passed"
		if runErr != nil {
			c.Status, c.Error = "failed", runErr.Error()
		}
		if err := r.save(); err != nil {
			return r, err
		}
	}
	r.Status = "passed"
	for _, c := range r.Cases {
		if c.Status != "passed" {
			r.Status = "failed"
		}
	}
	if err := r.save(); err != nil {
		return r, err
	}
	if r.Status != "passed" {
		return r, fmt.Errorf("corpus has failed or incomplete cases; see %s", r.SavedReport)
	}
	return r, nil
}

func corpusReportMatches(r *Report, artifactHash string, suite Suite, key string) error {
	if r.SHA256 != artifactHash || r.SuiteSHA256 != office.Hash(project.JSON(suite)) || r.ToolVersion != project.Version {
		return fmt.Errorf("artifact, suite or runner changed during corpus execution")
	}
	inputs, err := readInputs(suite.Inputs, r.InputSnapshots)
	if err != nil {
		return err
	}
	if corpusIdentity(artifactHash, suite, inputs) != key {
		return fmt.Errorf("input bytes changed during corpus execution")
	}
	return nil
}
