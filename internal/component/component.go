// Package component installs small, editable building blocks into WordUp workspaces.
package component

import (
	"bytes"
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/structure"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"runtime"
	"slices"
	"sort"
	"strings"
	"unicode/utf8"
)

const lockPath = ".wordwright/components.json"

// A local bundle is an editable source input, not an unbounded archive. Keep
// manifest intake bounded before reading all of its payloads so a cloned
// component cannot turn a fast install/check loop into a memory or I/O sink.
const maxBundleFiles = 4096

type File struct {
	Path   string `json:"path"`
	Text   string `json:"text,omitempty"`
	Binary bool   `json:"binary,omitempty"`
	data   []byte
}
type RibbonMerge struct {
	Source string `json:"source"`
	Target string `json:"target"`
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
	RibbonMerges       []RibbonMerge     `json:"ribbon_merges,omitempty"`
	applied            map[string]string
}
type Installed struct {
	ID                 string            `json:"id"`
	Version            string            `json:"version"`
	Files              map[string]string `json:"files"`
	Provenance         string            `json:"provenance,omitempty"`
	License            string            `json:"license,omitempty"`
	ManifestSHA256     string            `json:"manifest_sha256,omitempty"`
	Parameters         map[string]string `json:"parameters,omitempty"`
	SupportedPlatforms []string          `json:"supported_platforms,omitempty"`
	RibbonMerges       []RibbonMerge     `json:"ribbon_merges,omitempty"`
}
type Lock struct {
	Schema     int                  `json:"schema"`
	Components map[string]Installed `json:"components"`
}

type exportedSymbol struct {
	Name string
	Kind string
	Path string
	Line int
}

type manifestFileHash struct {
	Path   string `json:"path"`
	SHA256 string `json:"sha256"`
	Binary bool   `json:"binary,omitempty"`
}

// manifestHash binds provenance to the declared manifest and the exact bytes
// of every listed file. Binary payloads live outside Manifest's JSON source
// field, so hashing only project.JSON(m) would silently omit them.
func manifestHash(m Manifest) string {
	files := make([]manifestFileHash, 0, len(m.Files))
	for _, file := range m.Files {
		files = append(files, manifestFileHash{Path: file.Path, SHA256: office.Hash(fileData(file)), Binary: file.Binary})
	}
	m.Files = nil
	return office.Hash(project.JSON(struct {
		Manifest Manifest           `json:"manifest"`
		Files    []manifestFileHash `json:"files"`
	}{Manifest: m, Files: files}))
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
		vbaComponent("document.style-converter", "WordUpStyleConverter", "WU_ConvertStyle", "Converts one paragraph style in a single safe edit.", "style conversion", styleConverterSource),
		vbaComponent("document.text-operations", "WordUpTextOperations", "WU_ReplaceLiteral", "Performs bounded literal text replacements across selected Word stories or an exact Range without using Selection.", "text operations", textOperationsSource)}
	for i := range items {
		if items[i].ID == "command.hotkey" {
			items[i].Version = "1.0.5"
		}
		if items[i].ID == "operation.safe-edit" {
			items[i].Version = "1.0.2"
		}
		if items[i].ID == "command.context-menu" {
			items[i].Version = "1.0.4"
		}
		if items[i].ID == "ui.form-shell" {
			items[i].Version = "1.0.2"
		}
		if items[i].ID == "ui.progress-cancel" {
			items[i].Version = "1.0.1"
		}
		if items[i].ID == "ui.ribbon-command" {
			items[i].Version = "1.0.3"
		}
		if items[i].ID == "document.style-converter" {
			items[i].Version = "1.0.3"
		}
		if items[i].ID == "document.text-operations" {
			items[i].Version = "1.0.5"
			items[i].Acceptance = "WU_ReplaceLiteral and WU_ReplaceLiteralInRange are public, bounded, no-op aware, and compile without non-Office references."
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
	if len(m.Files) > maxBundleFiles {
		return Manifest{}, fmt.Errorf("component.json file count exceeds %d", maxBundleFiles)
	}
	seen := map[string]bool{}
	var total int64
	for i := range m.Files {
		if m.Files[i].Path == "" || m.Files[i].Text != "" || m.Files[i].Binary {
			return Manifest{}, fmt.Errorf("component.json file %d requires a path and must not embed source or binary data", i)
		}
		path := filepath.ToSlash(m.Files[i].Path)
		// Bundle paths are a cross-platform wire format. Do not silently
		// normalize a Windows spelling and then record a different source
		// identity; the manifest must already use slash-separated paths.
		if path != m.Files[i].Path || !fs.ValidPath(path) || path == "." || strings.ContainsAny(path, `\:`) {
			return Manifest{}, fmt.Errorf("component.json file %d requires a canonical relative path: %q", i, m.Files[i].Path)
		}
		key := strings.ToLower(path)
		if seen[key] {
			return Manifest{}, fmt.Errorf("component.json has duplicate path %q", m.Files[i].Path)
		}
		seen[key] = true
		data, readErr := readBundleData(dir, m.Files[i].Path)
		err = readErr
		if err != nil {
			return Manifest{}, fmt.Errorf("component file %s: %w", m.Files[i].Path, err)
		}
		if total > int64(office.Limit)-int64(len(data)) {
			return Manifest{}, fmt.Errorf("component bundle byte budget exceeded")
		}
		total += int64(len(data))
		if utf8.Valid(data) && !bytes.Contains(data, []byte{0}) {
			m.Files[i].Text = string(data)
		} else {
			if !binaryComponentPath(m.Files[i].Path) {
				return Manifest{}, fmt.Errorf("component file %s is binary but its source path requires editable text", m.Files[i].Path)
			}
			m.Files[i].Binary = true
			m.Files[i].data = data
		}
	}
	return m, nil
}

func readBundleData(root, path string) ([]byte, error) {
	return project.Read(root, path)
}

func fileData(file File) []byte {
	if file.Binary {
		return file.data
	}
	return []byte(file.Text)
}

func binaryComponentPath(name string) bool {
	name = strings.ToLower(filepath.ToSlash(name))
	if strings.HasPrefix(name, "vba/") || strings.HasPrefix(name, "forms/") {
		return false
	}
	switch strings.ToLower(filepath.Ext(name)) {
	case ".bas", ".cls", ".vba", ".xml", ".rels", ".json", ".md":
		return false
	default:
		return true
	}
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
				if strings.HasPrefix(strings.ToLower(filepath.ToSlash(m.Files[i].Path)), "vba/") {
					m.Files[i].Text = replaceVBAIdentifierPrefix(m.Files[i].Text, "WU_", value+"_")
				}
			}
			m.Acceptance = replaceVBAIdentifierPrefix(m.Acceptance, "WU_", value+"_")
		default:
			return Manifest{}, fmt.Errorf("component %s parameter %s has no typed adapter", m.ID, name)
		}
	}
	m.applied = values
	return m, nil
}

// replaceVBAIdentifierPrefix changes an identifier prefix without rewriting
// apostrophe comments or arbitrary prose. VBA strings are scanned as text too,
// because callbacks and macro names are commonly stored in string literals.
// The token boundary keeps names such as XWU_Name untouched.
func replaceVBAIdentifierPrefix(source, from, to string) string {
	if source == "" || from == "" || len(from) > len(source) {
		return source
	}
	var out strings.Builder
	out.Grow(len(source))
	inString, inComment, statementStart := false, false, true
	for i := 0; i < len(source); {
		if inComment {
			out.WriteByte(source[i])
			if source[i] == '\n' || source[i] == '\r' {
				inComment = false
				statementStart = true
			}
			i++
			continue
		}
		if inString {
			if source[i] == '"' {
				out.WriteByte(source[i])
				if i+1 < len(source) && source[i+1] == '"' {
					out.WriteByte(source[i+1])
					i += 2
					continue
				}
				inString = false
				statementStart = false
				i++
				continue
			}
			if hasIdentifierPrefix(source, i, from) {
				out.WriteString(to)
				i += len(from)
				continue
			}
			out.WriteByte(source[i])
			i++
			continue
		}
		if source[i] == '"' {
			inString = true
			out.WriteByte(source[i])
			i++
			continue
		}
		if statementStart && i+3 <= len(source) && strings.EqualFold(source[i:i+3], "Rem") && (i+3 == len(source) || vbaWhitespace(source[i+3])) {
			out.WriteString(source[i : i+3])
			i += 3
			inComment = true
			statementStart = false
			continue
		}
		if source[i] == '\'' {
			inComment = true
			out.WriteByte(source[i])
			statementStart = false
			i++
			continue
		}
		if source[i] == ':' {
			out.WriteByte(source[i])
			statementStart = true
			i++
			continue
		}
		if source[i] == '\n' || source[i] == '\r' {
			out.WriteByte(source[i])
			statementStart = true
			i++
			continue
		}
		if statementStart && vbaWhitespace(source[i]) {
			out.WriteByte(source[i])
			i++
			continue
		}
		if hasIdentifierPrefix(source, i, from) {
			out.WriteString(to)
			i += len(from)
			statementStart = false
			continue
		}
		out.WriteByte(source[i])
		statementStart = false
		i++
	}
	return out.String()
}

func hasIdentifierPrefix(source string, index int, prefix string) bool {
	if index > 0 && vbaIdentifierByte(source[index-1]) {
		return false
	}
	if index+len(prefix) >= len(source) || !strings.EqualFold(source[index:index+len(prefix)], prefix) {
		return false
	}
	return vbaIdentifierByte(source[index+len(prefix)])
}

func vbaIdentifierByte(value byte) bool {
	return value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value >= '0' && value <= '9' || value == '_'
}

func vbaWhitespace(value byte) bool {
	return value == ' ' || value == '\t'
}

func install(root string, m Manifest) (Installed, error) {
	id := m.ID
	if err := project.ValidateSourceDirectories(root); err != nil {
		return Installed{}, err
	}
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
			same = same && prior.Files[f.Path] == office.Hash(fileData(f))
		}
		if status["state"] == "clean" && prior.Version == m.Version && same && maps.Equal(prior.Parameters, m.applied) && slices.Equal(prior.RibbonMerges, m.RibbonMerges) {
			return prior, nil
		}
		return Installed{}, fmt.Errorf("component %s is already installed and differs from its recorded source; inspect component.status or component.diff", id)
	}
	installed := Installed{ID: id, Version: m.Version, Files: map[string]string{}, Provenance: m.Provenance, License: m.License, ManifestSHA256: manifestHash(m), Parameters: m.applied, SupportedPlatforms: append([]string(nil), m.SupportedPlatforms...), RibbonMerges: append([]RibbonMerge(nil), m.RibbonMerges...)}
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
		installed.Files[f.Path] = office.Hash(fileData(f))
	}
	if err := validateRibbonMerges(root, id, m); err != nil {
		return Installed{}, err
	}
	if err := validateBundleForms(id, m.Files); err != nil {
		return Installed{}, err
	}
	if err := rejectIdentifierCollisions(root, id, m.Files); err != nil {
		return Installed{}, err
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
		_, writeErr := output.Write(fileData(f))
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

// validateBundleForms performs the cheap structural checks that can be made
// without opening an existing MSForms binary. Build still owns full property
// and serialization validation; this preflight prevents an invalid bundle
// from partially installing before that later stage rejects it.
func validateBundleForms(componentID string, files []File) error {
	for _, file := range files {
		path := filepath.ToSlash(file.Path)
		if !strings.HasPrefix(strings.ToLower(path), "forms/") || !strings.EqualFold(filepath.Ext(path), ".json") {
			continue
		}
		var design office.Design
		if err := project.ReadJSON(fileData(file), &design); err != nil {
			return fmt.Errorf("component %s form %s: invalid design: %w", componentID, file.Path, err)
		}
		expected := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		if design.Name != expected {
			return fmt.Errorf("component %s form %s name %q does not match filename %q", componentID, file.Path, design.Name, expected)
		}
		if err := validateBundleControlSet(componentID, file.Path, "controls", design.Controls); err != nil {
			return err
		}
		if err := validateBundleControlSet(componentID, file.Path, "pages", design.Pages); err != nil {
			return err
		}
	}
	return nil
}

func validateBundleControlSet(componentID, filePath, scope string, controls []office.ControlDesign) error {
	seen := map[string]string{}
	for _, control := range controls {
		if !office.ValidIdentifier(control.Name) {
			return fmt.Errorf("component %s form %s has invalid control name %q in %s", componentID, filePath, control.Name, scope)
		}
		key := strings.ToLower(control.Name)
		if prior, exists := seen[key]; exists {
			return fmt.Errorf("component %s form %s has duplicate control name %q in %s (also declared as %q)", componentID, filePath, control.Name, scope, prior)
		}
		seen[key] = control.Name
		if err := validateBundleControlSet(componentID, filePath, control.Name+" controls", control.Controls); err != nil {
			return err
		}
		if err := validateBundleControlSet(componentID, filePath, control.Name+" pages", control.Pages); err != nil {
			return err
		}
	}
	return nil
}

func validateRibbonMerges(root, componentID string, m Manifest) error {
	if len(m.RibbonMerges) == 0 {
		return nil
	}
	files := map[string]File{}
	for _, file := range m.Files {
		files[strings.ToLower(filepath.ToSlash(file.Path))] = file
	}
	working := map[string][]byte{}
	loaded := map[string]bool{}
	hasBase := map[string]bool{}
	targetNames := map[string]string{}
	for _, merge := range m.RibbonMerges {
		sourceKey := strings.ToLower(filepath.ToSlash(merge.Source))
		file, ok := files[sourceKey]
		if !ok || !fs.ValidPath(merge.Source) || strings.ContainsAny(merge.Source, `\:`) || strings.HasPrefix(sourceKey, "package/") || strings.HasPrefix(sourceKey, ".wordwright/") {
			return fmt.Errorf("component %s Ribbon merge source %q must name a listed non-package source file", componentID, merge.Source)
		}
		target := filepath.ToSlash(merge.Target)
		if !office.SafePart(target) || !strings.HasSuffix(strings.ToLower(target), ".xml") {
			return fmt.Errorf("component %s Ribbon merge target %q must be a safe XML package part", componentID, merge.Target)
		}
		targetKey := strings.ToLower(target)
		if canonical, exists := targetNames[targetKey]; exists {
			target = canonical
		} else {
			targetNames[targetKey] = target
		}
		if !loaded[targetKey] {
			loaded[targetKey] = true
			if raw, err := project.Read(root, "package/"+target); err == nil {
				working[targetKey] = raw
				hasBase[targetKey] = true
			} else if !os.IsNotExist(err) {
				return err
			} else if baseRaw, baseErr := project.Read(root, ".wordwright/base.opc"); baseErr == nil {
				base, readErr := office.ReadPackage(baseRaw)
				if readErr != nil {
					return fmt.Errorf("component %s Ribbon base package: %w", componentID, readErr)
				}
				for name, raw := range base.Files {
					if strings.EqualFold(name, target) {
						working[targetKey] = append([]byte(nil), raw...)
						hasBase[targetKey] = true
						break
					}
				}
			} else if !os.IsNotExist(baseErr) {
				return baseErr
			}
		}
		if !hasBase[targetKey] {
			probe := office.BlankPackage()
			if err := probe.MergeRibbon(target, fileData(file)); err != nil {
				return fmt.Errorf("component %s Ribbon merge %s -> %s: %w", componentID, merge.Source, merge.Target, err)
			}
			working[targetKey] = append([]byte(nil), probe.Files[target]...)
			continue
		}
		merged, err := office.MergeRibbonXML(working[targetKey], fileData(file))
		if err != nil {
			return fmt.Errorf("component %s Ribbon merge %s -> %s: %w", componentID, merge.Source, merge.Target, err)
		}
		working[targetKey] = merged
	}
	return nil
}

func rejectIdentifierCollisions(root, componentID string, files []File) error {
	existing, err := workspaceExportedSymbols(root)
	if err != nil {
		return err
	}
	seen := map[string]exportedSymbol{}
	for _, file := range files {
		if !strings.HasPrefix(strings.ToLower(filepath.ToSlash(file.Path)), "vba/") {
			continue
		}
		base := filepath.Base(file.Path)
		ext := filepath.Ext(base)
		expectedModule := strings.TrimSuffix(base, ext)
		if !office.ValidIdentifier(expectedModule) {
			return fmt.Errorf("component %s source %s derives invalid VBA module name %q", componentID, file.Path, expectedModule)
		}
		declared, line, present := moduleNameAttribute(strings.Split(strings.ReplaceAll(file.Text, "\r\n", "\n"), "\n"))
		if present {
			if declared == "" || !office.ValidIdentifier(declared) {
				return fmt.Errorf("component %s has invalid VB_Name attribute in %s:%d", componentID, file.Path, line)
			}
			if !strings.EqualFold(declared, expectedModule) {
				return fmt.Errorf("component %s VB_Name %q in %s:%d does not match module filename %q", componentID, declared, file.Path, line, expectedModule)
			}
		}
		for _, symbol := range exportedSymbols(file.Path, file.Text) {
			key := strings.ToLower(symbol.Name)
			if prior, ok := seen[key]; ok && !compatiblePropertyAccessors(prior.Kind, symbol.Kind) {
				return fmt.Errorf("component %s exports duplicate %s %q in %s:%d and %s:%d", componentID, collisionLabel(prior, symbol), symbol.Name, prior.Path, prior.Line, symbol.Path, symbol.Line)
			}
			seen[key] = symbol
			if prior, ok := existing[key]; ok && !compatiblePropertyAccessors(prior.Kind, symbol.Kind) {
				return fmt.Errorf("component %s %s %q in %s:%d conflicts with existing %s:%d", componentID, collisionLabel(prior, symbol), symbol.Name, symbol.Path, symbol.Line, prior.Path, prior.Line)
			}
		}
	}
	return nil
}

func collisionLabel(first, second exportedSymbol) string {
	if first.Kind == "module" || second.Kind == "module" {
		return "module name"
	}
	return "identifier"
}

func workspaceExportedSymbols(root string) (map[string]exportedSymbol, error) {
	result := map[string]exportedSymbol{}
	base := filepath.Join(root, "vba")
	err := filepath.WalkDir(base, func(path string, entry fs.DirEntry, walkErr error) error {
		if os.IsNotExist(walkErr) {
			return nil
		}
		if walkErr != nil {
			return walkErr
		}
		if entry.Type()&os.ModeSymlink != 0 {
			return fmt.Errorf("source symlink is not accepted: %s", path)
		}
		if entry.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(entry.Name()))
		if ext != ".bas" && ext != ".cls" && ext != ".vba" {
			return nil
		}
		rel, err := filepath.Rel(root, path)
		if err != nil {
			return err
		}
		rel = filepath.ToSlash(rel)
		data, err := project.Read(root, rel)
		if err != nil {
			return err
		}
		for _, symbol := range exportedSymbols(rel, string(data)) {
			key := strings.ToLower(symbol.Name)
			if prior, ok := result[key]; !ok || symbol.Line < prior.Line {
				result[key] = symbol
			}
		}
		return nil
	})
	return result, err
}

func exportedSymbols(path, source string) []exportedSymbol {
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	result := []exportedSymbol{}
	if name, line := vbaModuleName(lines); name != "" {
		result = append(result, exportedSymbol{Name: name, Kind: "module", Path: path, Line: line})
	} else {
		// VBA.Rewrite derives a missing VB_Name from the source filename. Use
		// that same rule during component preflight so a path-different module
		// collision is rejected before installation rather than at build time.
		ext := strings.ToLower(filepath.Ext(path))
		if ext == ".bas" || ext == ".cls" || ext == ".vba" {
			name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
			if office.ValidIdentifier(name) {
				result = append(result, exportedSymbol{Name: name, Kind: "module", Path: path, Line: 1})
			}
		}
	}
	inProcedure := false
	for index, raw := range lines {
		line := strings.TrimSpace(vbaCodeLine(raw))
		if line == "" || strings.HasPrefix(line, "#") || strings.HasPrefix(strings.ToLower(line), "attribute ") {
			continue
		}
		lowerLine := strings.ToLower(line)
		if inProcedure {
			if strings.HasPrefix(lowerLine, "end sub") || strings.HasPrefix(lowerLine, "end function") || strings.HasPrefix(lowerLine, "end property") {
				inProcedure = false
			}
			continue
		}
		if strings.HasPrefix(lowerLine, "private sub ") || strings.HasPrefix(lowerLine, "private function ") || strings.HasPrefix(lowerLine, "private property ") || strings.HasPrefix(lowerLine, "private static sub ") || strings.HasPrefix(lowerLine, "private static function ") {
			inProcedure = !strings.Contains(lowerLine, "end sub") && !strings.Contains(lowerLine, "end function") && !strings.Contains(lowerLine, "end property")
			continue
		}
		fields := strings.Fields(line)
		if len(fields) == 0 {
			continue
		}
		position := 0
		explicitExport := false
		if strings.EqualFold(fields[position], "Private") {
			continue
		}
		if strings.EqualFold(fields[position], "Public") || strings.EqualFold(fields[position], "Friend") {
			explicitExport = true
			position++
		} else if strings.EqualFold(fields[position], "Static") {
			position++
		}
		if position >= len(fields) {
			continue
		}
		kind := strings.ToLower(fields[position])
		position++
		if kind == "property" {
			if position >= len(fields) {
				continue
			}
			kind += " " + strings.ToLower(fields[position])
			position++
		}
		if kind == "declare" && position < len(fields) {
			kind = strings.ToLower(fields[position])
			position++
		}
		procedure := false
		switch kind {
		case "sub", "function", "property get", "property let", "property set":
			procedure = true
		case "const", "dim", "enum", "type", "event":
		default:
			continue
		}
		if position >= len(fields) {
			continue
		}
		name := strings.TrimLeft(fields[position], "[")
		if cut := strings.IndexAny(name, "([%&^!#$@]"); cut >= 0 {
			name = name[:cut]
		}
		if name == "" || !office.ValidIdentifier(name) {
			continue
		}
		if !procedure && !explicitExport {
			continue
		}
		result = append(result, exportedSymbol{Name: name, Kind: kind, Path: path, Line: index + 1})
		if procedure {
			inProcedure = !strings.Contains(lowerLine, "end "+strings.Split(kind, " ")[0])
		}
	}
	return result
}

func vbaModuleName(lines []string) (string, int) {
	name, line, present := moduleNameAttribute(lines)
	if present && office.ValidIdentifier(name) {
		return name, line
	}
	return "", 0
}

// moduleNameAttribute finds the first VB_Name attribute and reports malformed
// or invalid values as present with an empty/invalid name. Callers that need
// to align with the builder must reject a present invalid attribute instead of
// silently deriving a second name from the source filename.
func moduleNameAttribute(lines []string) (string, int, bool) {
	for index, raw := range lines {
		line := strings.TrimSpace(raw)
		lower := strings.ToLower(line)
		const prefix = "attribute vb_name"
		if !strings.HasPrefix(lower, prefix) {
			continue
		}
		rest := strings.TrimSpace(line[len(prefix):])
		if !strings.HasPrefix(rest, "=") {
			continue
		}
		value := strings.TrimSpace(rest[1:])
		if len(value) < 2 || value[0] != '"' {
			return "", index + 1, true
		}
		close := strings.IndexByte(value[1:], '"')
		if close < 0 {
			return "", index + 1, true
		}
		return value[1 : close+1], index + 1, true
	}
	return "", 0, false
}

// ModuleName returns the first VBA VB_Name value and its Attribute line.
// present remains true for malformed or invalid values so source-only callers
// can report the error instead of silently deriving a filename-based name.
func ModuleName(source string) (string, int, bool) {
	return moduleNameAttribute(strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n"))
}

func vbaCodeLine(line string) string {
	quoted := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if quoted && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			quoted = !quoted
		case '\'':
			if !quoted {
				return line[:i]
			}
		}
	}
	trimmed := strings.TrimSpace(line)
	if len(trimmed) >= 4 && strings.EqualFold(trimmed[:3], "Rem") && vbaWhitespace(trimmed[3]) {
		return ""
	}
	return line
}

func compatiblePropertyAccessors(first, second string) bool {
	return strings.HasPrefix(first, "property ") && strings.HasPrefix(second, "property ") && first != second
}

func Status(root, id string) (map[string]any, error) {
	if err := project.ValidateSourceDirectories(root); err != nil {
		return nil, err
	}
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
	result := map[string]any{"id": id, "version": installed.Version, "state": state, "files": rows,
		"provenance": installed.Provenance, "license": installed.License, "manifest_sha256": installed.ManifestSHA256, "parameters": installed.Parameters,
		"supported_platforms": installed.SupportedPlatforms, "ribbon_merges": installed.RibbonMerges, "compatibility": Compatibility(installed.SupportedPlatforms)}
	// Built-ins have a local starting version to compare against. Local bundles
	// remain provenance-only after installation, so their current source cannot
	// be guessed from an ID alone.
	if bundled, getErr := Get(id); getErr == nil {
		adapted, adaptErr := adapt(bundled, installed.Parameters)
		if adaptErr == nil {
			bundleHash := manifestHash(adapted)
			result["bundled_version"] = adapted.Version
			result["bundled_manifest_sha256"] = bundleHash
			if installed.Version == adapted.Version && installed.ManifestSHA256 == bundleHash {
				result["version_state"] = "current"
			} else {
				result["version_state"] = "outdated"
			}
		} else {
			result["version_state"] = "unavailable"
			result["version_error"] = adaptErr.Error()
		}
	} else {
		result["version_state"] = "unknown"
	}
	return result, nil
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

// Compatibility reports whether the installed component can run on this host.
// An empty list means the source did not declare a platform, so status remains
// explicit rather than guessing.
func Compatibility(platforms []string) string {
	if len(platforms) == 0 {
		return "unknown"
	}
	for _, platform := range platforms {
		if platform == "*" || strings.EqualFold(platform, "all") || strings.EqualFold(platform, runtime.GOOS) {
			return "compatible"
		}
	}
	return "unsupported"
}

func diff(root string, m Manifest) (map[string]any, error) {
	status, err := Status(root, m.ID)
	if err != nil {
		return nil, err
	}
	files := make([]map[string]any, 0, len(m.Files))
	equal := true
	for _, f := range m.Files {
		bundled := fileData(f)
		row := map[string]any{"path": f.Path, "bundled_sha256": office.Hash(bundled), "bundled_bytes": len(bundled)}
		installed, readErr := project.Read(root, f.Path)
		switch {
		case os.IsNotExist(readErr):
			row["state"] = "missing"
			equal = false
		case readErr != nil:
			row["state"] = "unreadable"
			row["error"] = readErr.Error()
			equal = false
		default:
			row["installed_sha256"] = office.Hash(installed)
			row["installed_bytes"] = len(installed)
			if bytes.Equal(installed, bundled) {
				row["state"] = "clean"
			} else {
				row["state"] = "modified"
				row["first_difference_byte"] = firstDifference(bundled, installed)
				equal = false
			}
		}
		files = append(files, row)
	}
	return map[string]any{"component": m.ID, "bundled_version": m.Version, "status": status, "equal": equal, "files": files, "ribbon_merges": m.RibbonMerges}, nil
}

func firstDifference(a, b []byte) int {
	limit := len(a)
	if len(b) < limit {
		limit = len(b)
	}
	for i := 0; i < limit; i++ {
		if a[i] != b[i] {
			return i
		}
	}
	return limit
}
