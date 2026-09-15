package verify

import (
	"encoding/json"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

type frozenFile struct {
	Source string `json:"source"`
	Path   string `json:"path"`
	SHA256 string `json:"sha256"`
}
type frozenBundle struct {
	Schema       int          `json:"schema"`
	Kind         string       `json:"kind"`
	ReportSHA256 string       `json:"report_sha256"`
	Files        []frozenFile `json:"files"`
}

// Freeze retains source bytes; it never regenerates expected XML or changes
// the recorded suite. A bundle is local evidence, not a signature of truth.
func Freeze(reference, destination string) (map[string]any, error) {
	raw, err := project.Read(filepath.Dir(reference), filepath.Base(reference))
	if err != nil {
		return nil, err
	}
	var report Report
	if err := DecodeReport(raw, &report); err != nil {
		return nil, err
	}
	if err := checkedReport(&report); err != nil {
		return nil, err
	}
	if report.ArtifactSnapshot == "" || report.EvidenceDirectory == "" {
		return nil, fmt.Errorf("baseline lacks retained artifact/evidence; recapture it")
	}
	artifact, err := project.Read(filepath.Dir(report.ArtifactSnapshot), filepath.Base(report.ArtifactSnapshot))
	if err != nil {
		return nil, err
	}
	if office.Hash(artifact) != report.SHA256 {
		return nil, fmt.Errorf("baseline artifact snapshot hash mismatch")
	}
	destination, err = filepath.Abs(destination)
	if err != nil {
		return nil, err
	}
	if _, err := os.Lstat(destination); !os.IsNotExist(err) {
		return nil, fmt.Errorf("freeze destination must not exist")
	}
	parent := filepath.Dir(destination)
	if err := os.MkdirAll(parent, 0700); err != nil {
		return nil, err
	}
	staging, err := os.MkdirTemp(parent, ".wordup-freeze-")
	if err != nil {
		return nil, err
	}
	defer os.RemoveAll(staging) // Only this newly allocated staging directory.
	bundle := frozenBundle{Schema: 1, Kind: "wordup-parity", ReportSHA256: office.Hash(raw)}
	total := 0
	seen := map[string]bool{}
	copyFile := func(source, relative string) error {
		if seen[source] {
			return nil
		}
		seen[source] = true
		if len(bundle.Files) >= 4096 {
			return fmt.Errorf("freeze file limit exceeded")
		}
		data, err := project.Read(filepath.Dir(source), filepath.Base(source))
		if err != nil {
			return err
		}
		total += len(data)
		if total > 512<<20 {
			return fmt.Errorf("freeze byte budget exceeded")
		}
		if err := project.Write(staging, relative, data, ""); err != nil {
			return err
		}
		bundle.Files = append(bundle.Files, frozenFile{source, relative, office.Hash(data)})
		return nil
	}
	if err := copyFile(report.ArtifactSnapshot, "artifact"+filepath.Ext(report.ArtifactSnapshot)); err != nil {
		return nil, err
	}
	// Map iteration order is deliberately randomized. Keep frozen bundles
	// byte-stable so moving or regenerating the same parity fixture produces a
	// useful Git diff instead of noise.
	inputNames := make([]string, 0, len(report.InputSnapshots))
	for name := range report.InputSnapshots {
		inputNames = append(inputNames, name)
	}
	sort.Strings(inputNames)
	for _, name := range inputNames {
		input := report.InputSnapshots[name]
		if err := copyFile(input.File, "inputs/"+name+filepath.Ext(input.Source)); err != nil {
			return nil, err
		}
	}
	for i, step := range report.Suite.Steps {
		if step.Operation.Op != "xml.snapshot" {
			continue
		}
		if _, err := snapshotBytes(report.Observations[i].Result); err != nil {
			return nil, err
		}
		value, _ := canonical(report.Observations[i].Result).(map[string]any)
		if file, ok := value["file"].(string); ok && file != "" {
			if err := copyFile(file, fmt.Sprintf("snapshots/%04d.xml", i)); err != nil {
				return nil, err
			}
		}
	}
	err = filepath.WalkDir(report.EvidenceDirectory, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.Type()&os.ModeSymlink != 0 {
			return fmt.Errorf("evidence symlink rejected: %s", path)
		}
		if entry.IsDir() {
			return nil
		}
		rel, err := filepath.Rel(report.EvidenceDirectory, path)
		if err != nil {
			return err
		}
		return copyFile(path, "evidence/"+filepath.ToSlash(rel))
	})
	if err != nil {
		return nil, err
	}
	if err := project.Write(staging, "report.json", raw, ""); err != nil {
		return nil, err
	}
	if err := project.Write(staging, "bundle.json", project.JSON(bundle), ""); err != nil {
		return nil, err
	}
	// Validate the complete relocated view before exposing the directory.
	if _, _, err := LoadReport(filepath.Join(staging, "bundle.json")); err != nil {
		return nil, err
	}
	if err := os.Rename(staging, destination); err != nil {
		return nil, err
	}
	return map[string]any{"reference": filepath.Join(destination, "bundle.json"), "report_sha256": bundle.ReportSHA256, "files": len(bundle.Files), "bytes": total, "xml_rewritten": false}, nil
}

// LoadReport resolves only recorded evidence paths. Suite bytes and its hash
// remain unchanged; XML and document files are read and verified, never edited.
func LoadReport(path string) (*Report, string, error) {
	raw, err := project.Read(filepath.Dir(path), filepath.Base(path))
	if err != nil {
		return nil, "", err
	}
	var discriminator struct {
		Kind string `json:"kind"`
	}
	if err := json.Unmarshal(raw, &discriminator); err != nil {
		return nil, "", err
	}
	if discriminator.Kind != "wordup-parity" {
		var report Report
		err := DecodeReport(raw, &report)
		return &report, office.Hash(raw), err
	}
	var bundle frozenBundle
	if err := project.ReadJSON(raw, &bundle); err != nil {
		return nil, "", err
	}
	if bundle.Schema != 1 || len(bundle.Files) > 4096 {
		return nil, "", fmt.Errorf("invalid parity bundle")
	}
	root := filepath.Dir(path)
	reportBytes, err := project.Read(root, "report.json")
	if err != nil {
		return nil, "", err
	}
	if office.Hash(reportBytes) != bundle.ReportSHA256 {
		return nil, "", fmt.Errorf("frozen report hash mismatch")
	}
	var report Report
	if err := DecodeReport(reportBytes, &report); err != nil {
		return nil, "", err
	}
	paths := map[string]string{}
	destinations := map[string]bool{}
	total := 0
	for _, file := range bundle.Files {
		key := strings.ToLower(file.Path)
		if file.Source == "" || paths[file.Source] != "" || destinations[key] {
			return nil, "", fmt.Errorf("duplicate/empty bundle path")
		}
		destinations[key] = true
		data, err := project.Read(root, file.Path)
		if err != nil {
			return nil, "", err
		}
		total += len(data)
		if total > 512<<20 || office.Hash(data) != file.SHA256 {
			return nil, "", fmt.Errorf("frozen file hash/budget mismatch: %s", file.Path)
		}
		resolved, err := project.Under(root, file.Path)
		if err != nil {
			return nil, "", err
		}
		paths[file.Source] = resolved
	}
	report.ArtifactSnapshot = paths[report.ArtifactSnapshot]
	if report.ArtifactSnapshot == "" {
		return nil, "", fmt.Errorf("frozen artifact missing")
	}
	for name, input := range report.InputSnapshots {
		input.File = paths[input.File]
		if input.File == "" {
			return nil, "", fmt.Errorf("frozen input missing: %s", name)
		}
		report.InputSnapshots[name] = input
	}
	var relocate func(any) any
	relocate = func(value any) any {
		switch v := value.(type) {
		case map[string]any:
			for key, item := range v {
				if key == "file" || key == "path" || key == "staged_path" {
					if old, ok := item.(string); ok {
						if replacement, exists := paths[old]; exists {
							v[key] = replacement
							continue
						}
					}
				}
				v[key] = relocate(item)
			}
			return v
		case []any:
			for i, item := range v {
				v[i] = relocate(item)
			}
			return v
		}
		return value
	}
	for i := range report.Observations {
		ob := &report.Observations[i]
		ob.Result = relocate(ob.Result)
		ob.Diagnostics = relocate(ob.Diagnostics)
		ob.FailureXML = relocate(ob.FailureXML)
		if i < len(report.Suite.Steps) && report.Suite.Steps[i].Operation.Op == "xml.snapshot" {
			value, _ := canonical(ob.Result).(map[string]any)
			if file, ok := value["file"].(string); ok {
				rel, err := filepath.Rel(root, file)
				if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
					return nil, "", fmt.Errorf("snapshot outside frozen bundle")
				}
			}
			if _, err := snapshotBytes(ob.Result); err != nil {
				return nil, "", err
			}
		}
	}
	report.SavedReport = filepath.Join(root, "report.json")
	report.EvidenceDirectory = filepath.Join(root, "evidence")
	if err := checkedReport(&report); err != nil {
		return nil, "", err
	}
	artifact, err := project.Read(filepath.Dir(report.ArtifactSnapshot), filepath.Base(report.ArtifactSnapshot))
	if err != nil || office.Hash(artifact) != report.SHA256 {
		return nil, "", fmt.Errorf("frozen artifact identity mismatch: %v", err)
	}
	return &report, bundle.ReportSHA256, nil
}
