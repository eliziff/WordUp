package agent

import (
	"encoding/json"
	"fmt"
	"io/fs"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/project"
)

func analysisRequest(root string, paths []string) (string, error) {
	if len(paths) == 0 || len(paths) > 512 {
		return "", fmt.Errorf("vba.analyze requires 1..512 workspace paths")
	}
	w, err := project.Open(root)
	if err != nil {
		return "", err
	}
	modules := []map[string]string{}
	seen := map[string]bool{}
	total := 0
	for _, path := range paths {
		path = filepath.ToSlash(path)
		if !fs.ValidPath(path) || !strings.HasPrefix(path, "vba/") {
			return "", fmt.Errorf("analysis path must be under vba/: %s", path)
		}
		name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		if seen[strings.ToLower(name)] {
			return "", fmt.Errorf("duplicate analysis module: %s", name)
		}
		seen[strings.ToLower(name)] = true
		kind := w.Manifest.Components[name]
		if kind == "" {
			for declared, declaredKind := range w.Manifest.Components {
				if strings.EqualFold(declared, name) {
					kind = declaredKind
					break
				}
			}
		}
		if kind == "" {
			switch strings.ToLower(filepath.Ext(path)) {
			case ".bas":
				kind = "standard"
			case ".cls":
				kind = "class"
			}
		}
		if kind != "standard" && kind != "class" {
			return "", fmt.Errorf("analysis host/designer metadata not integrated for %s (%s)", path, kind)
		}
		data, err := project.Read(root, path)
		if err != nil {
			return "", err
		}
		total += len(data)
		if total > 4*1024*1024 {
			return "", fmt.Errorf("analysis source exceeds 4 MiB; select fewer modules")
		}
		if !utf8.Valid(data) {
			return "", fmt.Errorf("analysis requires UTF-8 source: %s", path)
		}
		modules = append(modules, map[string]string{"Name": name, "Kind": kind, "Path": path, "Source": string(data)})
	}
	data, err := json.Marshal(map[string]any{"Modules": modules, "Inspections": []string{"OptionExplicitInspection", "VariableNotUsedInspection"}})
	if err != nil {
		return "", err
	}
	// Include the outer helper array/escaping in the worker's request limit.
	wire, err := json.Marshal([]string{"analyze", string(data)})
	if err != nil {
		return "", err
	}
	if len(wire) > 4*1024*1024 {
		return "", fmt.Errorf("analysis request exceeds 4 MiB; select fewer modules")
	}
	return string(data), nil
}
