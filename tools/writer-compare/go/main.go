// Development-only comparison of the existing workspace build path.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
func run() error {
	if len(os.Args) < 3 {
		return fmt.Errorf("usage: writer-compare OUTPUT INPUT...")
	}
	root, err := filepath.Abs(os.Args[1])
	if err != nil {
		return err
	}
	if err := os.Mkdir(root, 0700); err != nil {
		return err
	}
	var rows []map[string]any
	for i, input := range os.Args[2:] {
		for _, operation := range []string{"unchanged", "module-edit", "class-add", "control-caption", "control-font"} {
			workspace := filepath.Join(root, fmt.Sprintf("%d-%s", i, operation))
			start := time.Now()
			if _, err := project.Import(input, workspace); err != nil {
				return err
			}
			importMS := float64(time.Since(start).Microseconds()) / 1000
			w, err := project.Open(workspace)
			if err != nil {
				return err
			}
			if operation == "module-edit" {
				files, err := filepath.Glob(filepath.Join(workspace, "vba", "*"))
				if err != nil || len(files) == 0 {
					return fmt.Errorf("no VBA files: %v", err)
				}
				sort.Strings(files)
				rel, _ := filepath.Rel(workspace, files[0])
				data, err := project.Read(workspace, filepath.ToSlash(rel))
				if err != nil {
					return err
				}
				data = []byte(strings.TrimRight(string(data), "\r\n") + "\r\n' WordUp writer comparison\r\n")
				if err := project.Write(workspace, filepath.ToSlash(rel), data, ""); err != nil {
					return err
				}
			}
			if operation == "class-add" {
				body := "Option Explicit\nPrivate value As Long\nPublic Property Get Current() As Long\nCurrent = value\nEnd Property\nPublic Property Let Current(ByVal nextValue As Long)\nvalue = nextValue\nEnd Property\n"
				if err := project.Write(workspace, "vba/WriterClass.cls", []byte(body), ""); err != nil {
					return err
				}
			}
			var control []string
			if operation == "control-caption" || operation == "control-font" {
				control, err = mutateForm(workspace, operation)
				if err != nil {
					return err
				}
			}
			var samples []float64
			var report *project.BuildReport
			for repetition := 0; repetition < 7; repetition++ {
				start = time.Now()
				// Different outputs prevent an artifact-cache hit: measure actual rebuilds.
				report, err = w.Build(filepath.Join(root, fmt.Sprintf("%d-%s-%d.dotm", i, operation, repetition)))
				if err != nil {
					return err
				}
				samples = append(samples, float64(time.Since(start).Microseconds())/1000)
			}
			warm := append([]float64(nil), samples[1:]...)
			sort.Float64s(warm)
			row := map[string]any{"input": input, "operation": operation, "import_ms": importMS, "samples_ms": samples,
				"warm_median_ms": (warm[2] + warm[3]) / 2, "output": report.Artifact, "cached": report.Cached,
				"native_word_executed": false, "scope": "workspace build timing; independent comparison required"}
			if control != nil {
				row["control"] = control
			}
			rows = append(rows, row)
		}
	}
	data, err := json.MarshalIndent(rows, "", "  ")
	if err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(root, "report.json"), data, 0600); err != nil {
		return err
	}
	fmt.Println(string(data))
	return nil
}

func mutateForm(root, operation string) ([]string, error) {
	files, err := filepath.Glob(filepath.Join(root, "forms", "*.json"))
	if err != nil {
		return nil, err
	}
	sort.Strings(files)
	type candidate struct {
		form, path string
		control    *office.ControlDesign
	}
	var choices []candidate
	var visit func(string, string, []office.ControlDesign)
	visit = func(form, parent string, controls []office.ControlDesign) {
		for i := range controls {
			control := &controls[i]
			path := control.Name
			if parent != "" {
				path = parent + "/" + path
			}
			if control.Type == "Label" || control.Type == "CommandButton" {
				choices = append(choices, candidate{form: form, path: path, control: control})
			}
			visit(form, path, control.Controls)
			visit(form, path, control.Pages)
		}
	}
	designs := map[string]*office.Design{}
	for _, file := range files {
		data, readErr := os.ReadFile(file)
		if readErr != nil {
			return nil, readErr
		}
		var design office.Design
		if readErr = json.Unmarshal(data, &design); readErr != nil {
			return nil, readErr
		}
		designs[file] = &design
		visit(design.Name, "", design.Controls)
	}
	if len(choices) == 0 {
		return nil, fmt.Errorf("precondition unmet: no supported caption control")
	}
	sort.Slice(choices, func(i, j int) bool {
		return choices[i].form+"\x00"+choices[i].path < choices[j].form+"\x00"+choices[j].path
	})
	selected := choices[0]
	if selected.control.Properties == nil {
		selected.control.Properties = map[string]any{}
	}
	if operation == "control-caption" {
		selected.control.Properties["Caption"] = "Writer comparison"
	} else {
		selected.control.Properties["Font"] = map[string]any{"Name": "Segoe UI", "Size": 12}
	}
	for file, design := range designs {
		if design.Name == selected.form {
			rel, relErr := filepath.Rel(root, file)
			if relErr != nil {
				return nil, relErr
			}
			if err := project.Write(root, filepath.ToSlash(rel), project.JSON(design), ""); err != nil {
				return nil, err
			}
			break
		}
	}
	return []string{selected.form, selected.control.Name}, nil
}
