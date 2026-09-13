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
		workspace := filepath.Join(root, fmt.Sprint(i))
		start := time.Now()
		if _, err := project.Import(input, workspace); err != nil {
			return err
		}
		importMS := float64(time.Since(start).Microseconds()) / 1000
		w, err := project.Open(workspace)
		if err != nil {
			return err
		}
		for _, operation := range []string{"unchanged", "module-edit"} {
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
			rows = append(rows, map[string]any{"input": input, "operation": operation, "import_ms": importMS, "samples_ms": samples,
				"warm_median_ms": (warm[2] + warm[3]) / 2, "output": report.Artifact, "cached": report.Cached,
				"native_word_executed": false, "scope": "workspace build timing; independent comparison required"})
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
