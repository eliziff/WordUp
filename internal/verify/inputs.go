package verify

import (
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"path/filepath"
	"regexp"
)

var inputName = regexp.MustCompile(`^[A-Za-z][A-Za-z0-9_]{0,63}$`)

type InputSnapshot struct {
	File   string `json:"file"`
	SHA256 string `json:"sha256"`
	Source string `json:"source"`
}
type inputBytes struct {
	source string
	data   []byte
}

func readInputs(declared map[string]string, frozen map[string]InputSnapshot) (map[string]inputBytes, error) {
	if frozen != nil && len(frozen) != len(declared) {
		return nil, fmt.Errorf("recorded input coverage differs; recapture the baseline")
	}
	result := map[string]inputBytes{}
	total := 0
	for name, source := range declared {
		path := source
		if frozen != nil {
			snapshot, ok := frozen[name]
			if !ok {
				return nil, fmt.Errorf("missing frozen input %s", name)
			}
			path = snapshot.File
		}
		data, err := project.Read(filepath.Dir(path), filepath.Base(path))
		if err != nil {
			return nil, fmt.Errorf("input %s: %w", name, err)
		}
		if frozen != nil && office.Hash(data) != frozen[name].SHA256 {
			return nil, fmt.Errorf("frozen input %s hash mismatch", name)
		}
		total += len(data)
		if total > 128<<20 {
			return nil, fmt.Errorf("suite input byte budget exceeded")
		}
		result[name] = inputBytes{source, data}
	}
	return result, nil
}

func stageInputs(output string, inputs map[string]inputBytes, replacements map[string]string) (map[string]InputSnapshot, error) {
	snapshots := map[string]InputSnapshot{}
	for name, input := range inputs {
		filename := name + filepath.Ext(input.source)
		snapshot := filepath.Join(output, "inputs", filename)
		// Native staging keys files by basename. Give each disposable test input
		// its own name so a prior SaveAs/Close cannot leave a stale basename that
		// blocks an unrelated run in the same warm host. Frozen names stay stable.
		working := filepath.Join(output, "working-inputs", name+"_"+office.Hash([]byte(output))[:16]+filepath.Ext(input.source))
		if err := project.AtomicWrite(snapshot, input.data); err != nil {
			return nil, err
		}
		if err := project.AtomicWrite(working, input.data); err != nil {
			return nil, err
		}
		snapshots[name] = InputSnapshot{File: snapshot, SHA256: office.Hash(input.data), Source: input.source}
		replacements["$input:"+name+"$"] = working
	}
	return snapshots, nil
}
