package gold

import (
	"bytes"
	"fmt"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

type Artifact struct {
	Path   string `json:"path"`
	SHA256 string `json:"sha256"`
}
type Counterexample struct {
	Name     string   `json:"name"`
	Artifact Artifact `json:"artifact"`
}
type Fixture struct {
	Schema          int              `json:"schema"`
	ID              string           `json:"id"`
	Author          string           `json:"author"`
	Status          string           `json:"status"`
	Operation       string           `json:"operation"`
	RevisionPolicy  string           `json:"revision_policy"`
	Preconditions   []string         `json:"preconditions"`
	Preserve        []string         `json:"preserve"`
	Input           Artifact         `json:"input"`
	Expected        Artifact         `json:"expected"`
	Comparison      string           `json:"comparison"`
	Counterexamples []Counterexample `json:"counterexamples"`
}

func artifact(root string, a Artifact) ([]byte, error) {
	data, err := project.Read(root, a.Path)
	if err != nil {
		return nil, err
	}
	if office.Hash(data) != a.SHA256 {
		return nil, fmt.Errorf("gold artifact hash mismatch: %s", a.Path)
	}
	if _, err := office.XMLSpans(data); err != nil {
		return nil, fmt.Errorf("gold artifact %s: %w", a.Path, err)
	}
	return data, nil
}

func equal(a, b []byte, mode string) (bool, error) {
	if mode == "exact_xml" {
		return bytes.Equal(a, b), nil
	}
	result, err := office.CompareXML(a, b, office.XMLComparePolicy{})
	if err != nil {
		return false, err
	}
	return result["equal"] == true, nil
}

// ValidateFixture proves identity, well-formedness and that named negative
// controls fail. It does not execute the operation or approve its specification.
func ValidateFixture(root, path string) (*Fixture, map[string]any, error) {
	var f Fixture
	raw, err := read(root, path, &f)
	if err != nil {
		return nil, nil, err
	}
	if f.Schema != 1 {
		return nil, nil, fmt.Errorf("gold fixture requires schema 1")
	}
	if f.Status == "" {
		f.Status = "proposed"
	}
	if f.Status != "proposed" && f.Status != "silver" {
		return nil, nil, fmt.Errorf("status must be proposed or silver; validation does not certify expectations")
	}
	if f.Comparison == "" {
		f.Comparison = "exact_xml"
	}
	if f.Comparison != "exact_xml" && f.Comparison != "semantic_xml" {
		return nil, nil, fmt.Errorf("comparison must be exact_xml or semantic_xml; no ignored attributes/elements")
	}
	if f.RevisionPolicy != "" && f.RevisionPolicy != "safe_setup" && f.RevisionPolicy != "tracked_suggestion" && f.RevisionPolicy != "direct_invoked" {
		return nil, nil, fmt.Errorf("unknown revision_policy")
	}
	if f.Input.Path != "" || f.Input.SHA256 != "" {
		if _, err := artifact(root, f.Input); err != nil {
			return nil, nil, err
		}
	}
	expected, err := artifact(root, f.Expected)
	if err != nil {
		return nil, nil, err
	}
	seen := map[string]bool{}
	for _, c := range f.Counterexamples {
		if c.Name == "" || seen[c.Name] {
			return nil, nil, fmt.Errorf("counterexample names must be nonempty and unique")
		}
		seen[c.Name] = true
		data, err := artifact(root, c.Artifact)
		if err != nil {
			return nil, nil, err
		}
		match, err := equal(expected, data, f.Comparison)
		if err != nil {
			return nil, nil, err
		}
		if match {
			return nil, nil, fmt.Errorf("counterexample %q incorrectly passes", c.Name)
		}
	}
	return &f, map[string]any{"valid": true, "fixture_sha256": office.Hash(raw), "status": f.Status, "id": f.ID, "negative_controls_rejected": len(f.Counterexamples), "editorially_approved": false, "operation_executed": false}, nil
}

func VerifyFixture(root, path, candidate string) (map[string]any, error) {
	f, report, err := ValidateFixture(root, path)
	if err != nil {
		return nil, err
	}
	expected, err := artifact(root, f.Expected)
	if err != nil {
		return nil, err
	}
	data, err := project.Read(root, candidate)
	if err != nil {
		return nil, err
	}
	// Even exact-byte verification rejects malformed candidate XML explicitly.
	if _, err := office.XMLSpans(data); err != nil {
		return nil, err
	}
	matched, err := equal(expected, data, f.Comparison)
	if err != nil {
		return nil, err
	}
	report["matches_expected"] = matched
	report["candidate_sha256"] = office.Hash(data)
	report["comparison"] = f.Comparison
	if !matched {
		return report, fmt.Errorf("candidate does not match proposed gold fixture %s", f.ID)
	}
	return report, nil
}
