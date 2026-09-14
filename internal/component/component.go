// Package component installs small, editable building blocks into WordUp workspaces.
package component

import (
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/structure"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

const lockPath = ".wordwright/components.json"

type File struct {
	Path string `json:"path"`
	Text string `json:"text,omitempty"`
}
type Manifest struct {
	Schema             int               `json:"schema"`
	ID                 string            `json:"id"`
	Version            string            `json:"version"`
	Description        string            `json:"description"`
	License            string            `json:"license"`
	Provenance         string            `json:"provenance"`
	SupportedPlatforms []string          `json:"supported_platforms"`
	Parameters         map[string]string `json:"parameters,omitempty"`
	Defaults           map[string]string `json:"defaults,omitempty"`
	Capabilities       []string          `json:"capabilities"`
	Acceptance         string            `json:"acceptance"`
	Adaptation         string            `json:"adaptation"`
	Files              []File            `json:"files"`
	applied            map[string]string
}
type Installed struct {
	ID             string            `json:"id"`
	Version        string            `json:"version"`
	Files          map[string]string `json:"files"`
	Provenance     string            `json:"provenance,omitempty"`
	License        string            `json:"license,omitempty"`
	ManifestSHA256 string            `json:"manifest_sha256,omitempty"`
	Parameters     map[string]string `json:"parameters,omitempty"`
}
type Lock struct {
	Schema     int                  `json:"schema"`
	Components map[string]Installed `json:"components"`
}

func builtin() []Manifest {
	items := []Manifest{{ID: "structure.detect", Version: structure.ContractVersion,
		Description:  "Editable, dependency-free Word VBA document structure detector.",
		Capabilities: []string{"paragraph evidence", "heading candidates", "resolved hierarchy", "ambiguity diagnostics"},
		Acceptance:   "WU_DetectStructure returns a read-only two-dimensional Variant table.", Adaptation: "Customize WU_CustomizeCandidate for project eligibility and role mapping.",
		Files: []File{{Path: "vba/WordUpStructure.bas", Text: structure.Source}}},
		vbaComponent("operation.safe-edit", "WordUpSafeEdit", "WU_BeginSafeEdit", "Paired begin/end procedures scope one undo record and restore ScreenUpdating inside the caller's error handler.", "document editing", safeEditSource),
		vbaComponent("ui.form-shell", "WordUpFormShell", "WU_ShowForm", "Shows a named UserForm through one stable entry point.", "UserForm lifecycle", formShellSource),
		vbaComponent("ui.progress-cancel", "WordUpProgress", "WU_CancelRequested", "Provides cooperative progress and cancellation checkpoints.", "progress and cancellation", progressSource),
		vbaComponent("ui.ribbon-command", "WordUpRibbon", "WU_RibbonCommand", "Provides a stable Ribbon callback dispatch seam.", "Ribbon callbacks", ribbonSource),
		vbaComponent("command.hotkey", "WordUpHotkey", "WU_RegisterHotkey", "Registers and removes a template-owned key binding.", "hotkeys", hotkeySource),
		vbaComponent("command.context-menu", "WordUpContextMenu", "WU_RegisterContextMenu", "Registers and removes a tagged context-menu command.", "context menus", contextMenuSource),
		vbaComponent("document.style-converter", "WordUpStyleConverter", "WU_ConvertStyle", "Converts one paragraph style in a single safe edit.", "style conversion", styleConverterSource)}
	for i := range items {
		if items[i].ID == "command.hotkey" {
			items[i].Version = "1.0.2"
		}
		if items[i].ID == "command.context-menu" {
			items[i].Version = "1.0.2"
		}
		if items[i].ID == "ui.progress-cancel" {
			items[i].Version = "1.0.1"
		}
		if items[i].ID == "ui.ribbon-command" {
			items[i].Version = "1.0.1"
		}
		if items[i].ID == "document.style-converter" {
			items[i].Version = "1.0.2"
		}
		items[i].Schema = 1
		items[i].License = "MIT"
		items[i].Provenance = "WordUp 0.4.0 bundled editable source"
		items[i].SupportedPlatforms = []string{"windows"}
		if items[i].Parameters == nil {
			items[i].Parameters = map[string]string{"module_prefix": "valid VBA identifier prefix"}
		}
		if items[i].Defaults == nil {
			items[i].Defaults = map[string]string{"module_prefix": "WU"}
		}
	}
	return items
}

func vbaComponent(id, module, entry, description, capability, source string) Manifest {
	return Manifest{ID: id, Version: "1.0.0", Description: description,
		Capabilities: []string{capability}, Acceptance: entry + " is public and the module compiles without non-Office references.",
		Adaptation: "Rename only the documented WU_ public entry points and keep cleanup paths intact.",
		Files:      []File{{Path: "vba/" + module + ".bas", Text: source}}}
}

func List() []Manifest {
	items := builtin()
	for i := range items {
		items[i].Files = nil
	}
	return items
}
func Get(id string) (Manifest, error) {
	for _, item := range builtin() {
		if item.ID == id {
			return item, nil
		}
	}
	return Manifest{}, fmt.Errorf("unknown component %q", id)
}

func readLock(root string) (Lock, error) {
	lock := Lock{Schema: 1, Components: map[string]Installed{}}
	b, err := project.Read(root, lockPath)
	if os.IsNotExist(err) {
		return lock, nil
	}
	if err != nil {
		return lock, err
	}
	if err = project.ReadJSON(b, &lock); err != nil {
		return lock, err
	}
	if lock.Schema != 1 || lock.Components == nil {
		return lock, fmt.Errorf("component lock schema mismatch")
	}
	return lock, nil
}

func Add(root, id string) (Installed, error) {
	return AddWith(root, id, nil)
}

func AddWith(root, id string, values map[string]string) (Installed, error) {
	m, err := Get(id)
	if err != nil {
		return Installed{}, err
	}
	m, err = adapt(m, values)
	if err != nil {
		return Installed{}, err
	}
	return install(root, m)
}

func LoadBundle(dir string) (Manifest, error) {
	raw, err := project.Read(dir, "component.json")
	if err != nil {
		return Manifest{}, fmt.Errorf("read component bundle manifest: %w", err)
	}
	var m Manifest
	if err = project.ReadJSON(raw, &m); err != nil {
		return Manifest{}, fmt.Errorf("component.json: %w", err)
	}
	if m.Schema != 1 || m.ID == "" || m.Version == "" || m.License == "" || m.Provenance == "" || len(m.Files) == 0 {
		return Manifest{}, fmt.Errorf("component.json requires schema 1, id, version, license, provenance, and files")
	}
	for i := range m.Files {
		if m.Files[i].Path == "" || m.Files[i].Text != "" {
			return Manifest{}, fmt.Errorf("component.json file %d requires path and must not embed text", i)
		}
		m.Files[i].Text, err = readBundleText(dir, m.Files[i].Path)
		if err != nil {
			return Manifest{}, fmt.Errorf("component file %s: %w", m.Files[i].Path, err)
		}
	}
	return m, nil
}

func readBundleText(root, path string) (string, error) {
	b, err := project.Read(root, path)
	if err != nil {
		return "", err
	}
	if strings.IndexByte(string(b), 0) >= 0 {
		return "", fmt.Errorf("bundled editable source contains NUL bytes")
	}
	return string(b), nil
}

func AddBundle(root, dir string) (Installed, error) {
	return AddBundleWith(root, dir, nil)
}

func AddBundleWith(root, dir string, values map[string]string) (Installed, error) {
	m, err := LoadBundle(dir)
	if err != nil {
		return Installed{}, err
	}
	m, err = adapt(m, values)
	if err != nil {
		return Installed{}, err
	}
	return install(root, m)
}

func DiffBundle(root, dir string) (map[string]any, error) {
	return DiffBundleWith(root, dir, nil)
}

func DiffBundleWith(root, dir string, values map[string]string) (map[string]any, error) {
	m, err := LoadBundle(dir)
	if err != nil {
		return nil, err
	}
	values, err = diffParameters(root, m.ID, values)
	if err != nil {
		return nil, err
	}
	m, err = adapt(m, values)
	if err != nil {
		return nil, err
	}
	return diff(root, m)
}

func adapt(m Manifest, supplied map[string]string) (Manifest, error) {
	values := map[string]string{}
	for name, value := range m.Defaults {
		values[name] = value
	}
	for name, value := range supplied {
		if _, declared := m.Parameters[name]; !declared {
			return Manifest{}, fmt.Errorf("component %s does not declare parameter %s", m.ID, name)
		}
		values[name] = value
	}
	for name := range m.Parameters {
		if _, ok := values[name]; !ok {
			return Manifest{}, fmt.Errorf("component %s requires parameter %s", m.ID, name)
		}
	}
	for name, value := range values {
		switch name {
		case "module_prefix":
			if !office.ValidIdentifier(value) {
				return Manifest{}, fmt.Errorf("component %s parameter module_prefix must be a valid VBA identifier", m.ID)
			}
			for i := range m.Files {
				m.Files[i].Text = strings.ReplaceAll(m.Files[i].Text, "WU_", value+"_")
			}
			m.Acceptance = strings.ReplaceAll(m.Acceptance, "WU_", value+"_")
		default:
			return Manifest{}, fmt.Errorf("component %s parameter %s has no typed adapter", m.ID, name)
		}
	}
	m.applied = values
	return m, nil
}

func install(root string, m Manifest) (Installed, error) {
	id := m.ID
	guard, err := project.Under(root, ".wordwright/component-install.lock")
	if err != nil {
		return Installed{}, err
	}
	if err = os.MkdirAll(filepath.Dir(guard), 0700); err != nil {
		return Installed{}, err
	}
	lease, err := os.OpenFile(guard, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0600)
	if err != nil {
		return Installed{}, fmt.Errorf("component installation lock (another installer or interrupted installation): %w", err)
	}
	defer func() { lease.Close(); os.Remove(guard) }()
	lock, err := readLock(root)
	if err != nil {
		return Installed{}, err
	}
	if prior, ok := lock.Components[id]; ok {
		status, err := Status(root, id)
		if err != nil {
			return Installed{}, err
		}
		same := len(prior.Files) == len(m.Files)
		for _, f := range m.Files {
			same = same && prior.Files[f.Path] == office.Hash([]byte(f.Text))
		}
		if status["state"] == "clean" && prior.Version == m.Version && same && maps.Equal(prior.Parameters, m.applied) {
			return prior, nil
		}
		return Installed{}, fmt.Errorf("component %s is already installed and differs from its recorded source; inspect component.status or component.diff", id)
	}
	installed := Installed{ID: id, Version: m.Version, Files: map[string]string{}, Provenance: m.Provenance, License: m.License, ManifestSHA256: office.Hash(project.JSON(m)), Parameters: m.applied}
	// Preflight every file before writing any source. Components are independent
	// copies: no file in this path is allowed to replace an existing file.
	seen := map[string]bool{}
	for _, f := range m.Files {
		if !fs.ValidPath(f.Path) || f.Path == "." || strings.ContainsAny(f.Path, `\:`) {
			return Installed{}, fmt.Errorf("component %s requires a canonical relative file path: %q", id, f.Path)
		}
		key := strings.ToLower(filepath.ToSlash(f.Path))
		if key == ".wordwright" || strings.HasPrefix(key, ".wordwright/") || key == ".git" || strings.HasPrefix(key, ".git/") || seen[key] {
			return Installed{}, fmt.Errorf("component %s has reserved or duplicate path %s", id, f.Path)
		}
		seen[key] = true
		if _, err = project.Read(root, f.Path); err == nil {
			return Installed{}, fmt.Errorf("component %s would overwrite %s", id, f.Path)
		} else if !os.IsNotExist(err) {
			return Installed{}, err
		}
		installed.Files[f.Path] = office.Hash([]byte(f.Text))
	}
	created := []string{}
	rollback := func(cause error) (Installed, error) {
		for i := len(created) - 1; i >= 0; i-- {
			path := created[i]
			b, readErr := project.Read(root, path)
			if readErr != nil || office.Hash(b) != installed.Files[path] {
				cause = fmt.Errorf("%w; rollback preserved changed or unreadable file %s", cause, path)
				continue
			}
			full, pathErr := project.Under(root, path)
			if pathErr != nil {
				cause = fmt.Errorf("%w; rollback path %s: %v", cause, path, pathErr)
			} else if removeErr := os.Remove(full); removeErr != nil {
				cause = fmt.Errorf("%w; rollback %s: %v", cause, path, removeErr)
			}
		}
		return Installed{}, cause
	}
	for _, f := range m.Files {
		full, pathErr := project.Under(root, f.Path)
		if pathErr != nil {
			return rollback(pathErr)
		}
		if err = os.MkdirAll(filepath.Dir(full), 0700); err != nil {
			return rollback(err)
		}
		// Exclusive creation also protects files created after preflight.
		output, createErr := os.OpenFile(full, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0600)
		if createErr != nil {
			return rollback(createErr)
		}
		_, writeErr := output.WriteString(f.Text)
		if writeErr == nil {
			writeErr = output.Sync()
		}
		closeErr := output.Close()
		if writeErr != nil || closeErr != nil {
			// This file is ours, but may be incomplete. Report it rather than
			// deleting an unverified partial file another editor could have changed.
			return rollback(fmt.Errorf("component write %s incomplete (inspect before retry): %v %v", f.Path, writeErr, closeErr))
		}
		created = append(created, f.Path)
	}
	lock.Components[id] = installed
	if err = project.Write(root, lockPath, project.JSON(lock), ""); err != nil {
		return rollback(err)
	}
	return installed, nil
}

func Status(root, id string) (map[string]any, error) {
	lock, err := readLock(root)
	if err != nil {
		return nil, err
	}
	installed, ok := lock.Components[id]
	if !ok {
		return map[string]any{"id": id, "state": "not_installed"}, nil
	}
	rows := []map[string]any{}
	state := "clean"
	names := make([]string, 0, len(installed.Files))
	for name := range installed.Files {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		row := map[string]any{"path": name, "installed_sha256": installed.Files[name]}
		b, readErr := project.Read(root, name)
		if readErr != nil {
			if os.IsNotExist(readErr) {
				row["state"] = "missing"
			} else {
				row["state"] = "unreadable"
				row["error"] = readErr.Error()
			}
			state = "modified"
		} else {
			row["current_sha256"] = office.Hash(b)
			if row["current_sha256"] == row["installed_sha256"] {
				row["state"] = "clean"
			} else {
				row["state"] = "modified"
				state = "modified"
			}
		}
		rows = append(rows, row)
	}
	return map[string]any{"id": id, "version": installed.Version, "state": state, "files": rows,
		"provenance": installed.Provenance, "license": installed.License, "manifest_sha256": installed.ManifestSHA256, "parameters": installed.Parameters}, nil
}

func Diff(root, id string) (map[string]any, error) {
	return DiffWith(root, id, nil)
}

func DiffWith(root, id string, values map[string]string) (map[string]any, error) {
	m, err := Get(id)
	if err != nil {
		return nil, err
	}
	values, err = diffParameters(root, m.ID, values)
	if err != nil {
		return nil, err
	}
	m, err = adapt(m, values)
	if err != nil {
		return nil, err
	}
	return diff(root, m)
}

func diffParameters(root, id string, supplied map[string]string) (map[string]string, error) {
	if supplied != nil {
		return supplied, nil
	}
	lock, err := readLock(root)
	if err != nil {
		return nil, err
	}
	if installed, ok := lock.Components[id]; ok {
		return installed.Parameters, nil
	}
	return nil, nil
}

func diff(root string, m Manifest) (map[string]any, error) {
	status, err := Status(root, m.ID)
	if err != nil {
		return nil, err
	}
	base := map[string]any{}
	for _, f := range m.Files {
		base[f.Path] = map[string]any{"sha256": office.Hash([]byte(f.Text)), "text": f.Text}
	}
	return map[string]any{"component": m.ID, "bundled_version": m.Version, "status": status, "bundled_files": base}, nil
}
