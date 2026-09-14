// Package project turns Office artifacts into ordinary, guarded source trees.
package project

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/eliziff/WordUp/internal/office"
)

const Version = "0.4.0"

type Manifest struct {
	Name           string      `json:"name"`
	References     []Reference `json:"add_references,omitempty"`
	DropSignatures bool        `json:"drop_signatures,omitempty"`
}
type Reference struct {
	Name        string `json:"name"`
	GUID        string `json:"guid"`
	Version     string `json:"version"`
	Path        string `json:"path,omitempty"`
	Description string `json:"description,omitempty"`
}
type Index struct {
	BaselineSHA256 string            `json:"baseline_sha256"`
	Files          map[string]string `json:"files"`
	Components     map[string]string `json:"components"`
}
type BuildReport struct {
	ToolSHA256        string   `json:"tool_sha256"`
	Schema            int      `json:"schema"`
	ToolVersion       string   `json:"tool_version"`
	Artifact          string   `json:"artifact"`
	SHA256            string   `json:"sha256"`
	SourceFingerprint string   `json:"source_fingerprint"`
	Bytes             int      `json:"bytes"`
	NoChange          bool     `json:"byte_identical_to_baseline"`
	Cached            bool     `json:"cached"`
	Modules           int      `json:"modules"`
	Forms             int      `json:"forms"`
	ModifiedParts     []string `json:"modified_parts"`
	DurationMS        float64  `json:"duration_ms"`
	PackageValidated  bool     `json:"package_validated"`
	VBACompiled       bool     `json:"vba_compiled"`
	WordExecuted      bool     `json:"word_executed"`
	MacExecuted       bool     `json:"mac_executed"`
	Warnings          []string `json:"warnings,omitempty"`
}
type Workspace struct {
	Root           string
	Manifest       Manifest
	Index          Index
	Baseline       *office.Package
	buildMemo      *buildMemo
	baselineVBA    *office.VBA
	baselineHashes map[string]string
	sourceMemo     map[string][]byte
	sourceStamp    map[string]fileStamp
}

type fileStamp struct {
	Size       int64
	ModifiedNS int64
	ChangedNS  int64
	Hash       string
}

type buildMemo struct {
	Output   string
	Sources  map[string]fileStamp
	Artifact fileStamp
	Evidence fileStamp
	Report   BuildReport
}

func JSON(v any) []byte { b, _ := json.MarshalIndent(v, "", "  "); return append(b, '\n') }
func ReadJSON(b []byte, v any) error {
	// Windows editors and PowerShell commonly prefix UTF-8 files with a BOM.
	b = bytes.TrimPrefix(b, []byte{0xef, 0xbb, 0xbf})
	d := json.NewDecoder(bytes.NewReader(b))
	d.DisallowUnknownFields()
	if e := d.Decode(v); e != nil {
		return e
	}
	var extra any
	if e := d.Decode(&extra); e != io.EOF {
		return fmt.Errorf("trailing JSON data")
	}
	return nil
}
func Under(root, rel string) (string, error) {
	rel = filepath.ToSlash(rel)
	if !office.SafePart(rel) {
		return "", fmt.Errorf("unsafe workspace path %q", rel)
	}
	root, e := filepath.Abs(root)
	if e != nil {
		return "", e
	}
	cur := root
	// A supplied root may be an intentional symlink; resolve that boundary once.
	if r, e := filepath.EvalSymlinks(root); e == nil {
		root, cur = r, r
	}
	for _, s := range strings.Split(rel, "/") {
		cur = filepath.Join(cur, s)
		i, e := os.Lstat(cur)
		if e == nil && i.Mode()&os.ModeSymlink != 0 {
			return "", fmt.Errorf("symlink inside workspace: %s", rel)
		}
		if e != nil && !os.IsNotExist(e) {
			return "", e
		}
	}
	return filepath.Join(root, filepath.FromSlash(rel)), nil
}
func Read(root, rel string) ([]byte, error) {
	p, e := Under(root, rel)
	if e != nil {
		return nil, e
	}
	f, e := os.Open(p)
	if e != nil {
		return nil, e
	}
	defer f.Close()
	i, e := f.Stat()
	if e != nil {
		return nil, e
	}
	if !i.Mode().IsRegular() || i.Size() > office.Limit {
		return nil, fmt.Errorf("not a bounded ordinary source file")
	}
	b, e := io.ReadAll(io.LimitReader(f, office.Limit+1))
	if len(b) > office.Limit {
		return nil, fmt.Errorf("file budget exceeded")
	}
	return b, e
}
func Write(root, rel string, b []byte, expected string) error {
	p, e := Under(root, rel)
	if e != nil {
		return e
	}
	if len(b) > office.Limit {
		return fmt.Errorf("file budget exceeded")
	}
	if expected != "" {
		old, e := Read(root, rel)
		if e != nil {
			return e
		}
		if office.Hash(old) != expected {
			return fmt.Errorf("stale source hash: reread before editing")
		}
	}
	if e = os.MkdirAll(filepath.Dir(p), 0700); e != nil {
		return e
	}
	return AtomicWrite(p, b)
}
func AtomicWrite(file string, b []byte) error {
	dir := filepath.Dir(file)
	if e := os.MkdirAll(dir, 0700); e != nil {
		return e
	}
	f, e := os.CreateTemp(dir, ".wordwright-*")
	if e != nil {
		return e
	}
	tmp := f.Name()
	defer os.Remove(tmp)
	if e = f.Chmod(0600); e != nil {
		f.Close()
		return e
	}
	if _, e = f.Write(b); e != nil {
		f.Close()
		return e
	}
	if e = f.Sync(); e != nil {
		f.Close()
		return e
	}
	if e = f.Close(); e != nil {
		return e
	}
	return replaceFile(tmp, file)
}
func Open(root string) (*Workspace, error) {
	r, e := filepath.Abs(root)
	if e != nil {
		return nil, e
	}
	w := &Workspace{Root: r}
	b, e := Read(root, "project.json")
	if e != nil {
		return nil, e
	}
	if e = ReadJSON(b, &w.Manifest); e != nil {
		return nil, e
	}
	b, e = Read(root, ".wordwright/index.json")
	if e != nil {
		return nil, e
	}
	if e = ReadJSON(b, &w.Index); e != nil {
		return nil, e
	}
	b, e = Read(root, ".wordwright/base.opc")
	if e != nil {
		return nil, e
	}
	if office.Hash(b) != w.Index.BaselineSHA256 {
		return nil, fmt.Errorf("baseline hash mismatch")
	}
	w.Baseline, e = office.ReadPackage(b)
	return w, e
}

// ModuleKind derives source kind from its extension and imported VBA metadata.
func (w *Workspace) ModuleKind(name, ext string) (string, error) {
	kind := map[string]string{".bas": "standard", ".cls": "class", ".vba": "form"}[strings.ToLower(ext)]
	if kind == "" {
		return "", fmt.Errorf("unrecognized module extension %s", ext)
	}
	for component, importedKind := range w.Index.Components {
		if strings.EqualFold(component, name) {
			return importedKind, nil
		}
	}
	return kind, nil
}

func Import(source, destination string) (map[string]any, error) {
	source, e := filepath.Abs(source)
	if e != nil {
		return nil, e
	}
	b, e := Read(filepath.Dir(source), filepath.Base(source))
	if e != nil {
		return nil, e
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return nil, e
	}
	if e = p.Validate(); e != nil {
		return nil, e
	}
	return importPackage(p, filepath.Base(source), destination)
}
func New(name, destination string) (map[string]any, error) {
	if !office.ValidIdentifier(name) {
		return nil, fmt.Errorf("use a VBA-safe project name")
	}
	p := office.BlankPackage()
	b, e := p.Bytes()
	if e != nil {
		return nil, e
	}
	p, e = office.ReadPackage(b)
	if e != nil {
		return nil, e
	}
	report, e := importPackage(p, name+".dotm", destination)
	if e != nil {
		return nil, e
	}
	w, e := Open(destination)
	if e != nil {
		return nil, e
	}
	w.Manifest.Name = name
	if e = Write(destination, "project.json", JSON(w.Manifest), ""); e != nil {
		return nil, e
	}
	return report, nil
}
func importPackage(p *office.Package, sourceName, destination string) (map[string]any, error) {
	if _, e := os.Lstat(destination); e == nil {
		return nil, fmt.Errorf("destination already exists; import will not overwrite a workspace")
	}
	if e := os.MkdirAll(destination, 0700); e != nil {
		return nil, e
	}
	good := false
	defer func() {
		if !good {
			os.RemoveAll(destination)
		}
	}()
	m := Manifest{Name: "WordProject"}
	components := map[string]string{}
	idx := Index{BaselineSHA256: office.Hash(p.Original), Files: map[string]string{}, Components: components}
	put := func(rel string, b []byte) error {
		idx.Files[rel] = office.Hash(b)
		return Write(destination, rel, b, "")
	}
	warnings := []string{}
	for name, b := range p.Files {
		if e := put("package/"+name, b); e != nil {
			return nil, e
		}
	}
	if vb := p.Files["word/vbaProject.bin"]; len(vb) > 0 {
		v, e := office.ReadVBA(vb)
		if e != nil {
			return nil, e
		}
		m.Name = v.Name
		for _, mod := range v.Modules {
			ext := map[string]string{"standard": ".bas", "class": ".cls", "document": ".cls", "form": ".vba"}[mod.Kind]
			rel := "vba/" + mod.Name + ext
			if e = put(rel, []byte(mod.Source)); e != nil {
				return nil, e
			}
			components[mod.Name] = mod.Kind
			if mod.Kind == "form" {
				form, e := office.ReadForm(v.CFB, mod.Name, v.Codepage)
				var data []byte
				if e != nil {
					data = JSON(office.Design{Name: mod.Name, Mode: "preserve"})
					warnings = append(warnings, mod.Name+": layout preserved opaque; "+e.Error())
				} else {
					data = JSON(form.Design())
				}
				if e = put("forms/"+mod.Name+".json", data); e != nil {
					return nil, e
				}
			}
		}
	} else {
		components["ThisDocument"] = "document"
		if e := put("vba/ThisDocument.cls", []byte(DocumentSource())); e != nil {
			return nil, e
		}
	}
	for _, d := range []string{"assets", "tests", "forms", "references", "reports", "dist"} {
		if e := os.MkdirAll(filepath.Join(destination, d), 0700); e != nil {
			return nil, e
		}
	}
	if e := Write(destination, ".wordwright/base.opc", p.Original, ""); e != nil {
		return nil, e
	}
	if e := Write(destination, "project.json", JSON(m), ""); e != nil {
		return nil, e
	}
	if e := Write(destination, ".wordwright/index.json", JSON(idx), ""); e != nil {
		return nil, e
	}
	if e := Write(destination, "AGENTS.md", []byte(AgentInstructions), ""); e != nil {
		return nil, e
	}
	if e := Write(destination, ".gitattributes", []byte(workspaceGitAttributes), ""); e != nil {
		return nil, e
	}
	if e := Write(destination, ".gitignore", []byte(workspaceGitIgnore), ""); e != nil {
		return nil, e
	}
	catalog := office.Catalog(p)
	if e := Write(destination, "reports/import.json", JSON(map[string]any{"source_name": sourceName, "source_sha256": idx.BaselineSHA256, "components": components, "catalog": catalog, "warnings": warnings, "word_executed": false}), ""); e != nil {
		return nil, e
	}
	good = true
	return map[string]any{"workspace": destination, "components": components, "catalog": catalog, "warnings": warnings, "word_executed": false}, nil
}

// Imported bytes and exact XML expectations must survive any user's autocrlf
// setting. Disabling conversion does not disable text diffs for source files.
const workspaceGitAttributes = `* -text
*.bas diff
*.cls diff
*.vba diff
*.xml diff
*.rels diff
*.json diff
*.md diff
*.opc binary
*.bin binary
*.dotm binary
*.docm binary
*.docx binary
*.png binary
*.jpg binary
*.emf binary
*.wmf binary
`

const workspaceGitIgnore = `/dist/
/reports/
/.wordwright/session.json
/.wordwright/session.log
/.wordwright/session.starting
/.wordwright/component-install.lock
/.wordwright/deploy/
`

func DocumentSource() string {
	return `Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "0{00020906-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
`
}

var sourceDirectories = []string{"package", "vba", "forms", "assets"}

const componentLockSource = ".wordwright/components.json"

func (w *Workspace) SourceFiles() (map[string][]byte, error) {
	out := map[string][]byte{}
	total := 0
	for _, top := range sourceDirectories {
		base := filepath.Join(w.Root, top)
		e := filepath.WalkDir(base, func(p string, d fs.DirEntry, e error) error {
			if os.IsNotExist(e) {
				return nil
			}
			if e != nil {
				return e
			}
			if d.Type()&os.ModeSymlink != 0 {
				return fmt.Errorf("source symlink is not accepted")
			}
			if d.IsDir() {
				return nil
			}
			rel, e := filepath.Rel(w.Root, p)
			if e != nil {
				return e
			}
			rel = filepath.ToSlash(rel)
			b, e := Read(w.Root, rel)
			if e != nil {
				return e
			}
			total += len(b)
			if total > office.Limit {
				return fmt.Errorf("workspace source budget exceeded")
			}
			out[rel] = b
			return nil
		})
		if e != nil {
			return nil, e
		}
	}
	b, e := Read(w.Root, "project.json")
	if e != nil {
		return nil, e
	}
	out["project.json"] = b
	return out, nil
}

// SourceFilesCached returns the current source snapshot while reusing bytes
// whose file stamps have not changed. It shares the same stamp/memo path as
// Build, so resident agents do not pay a full read for every check request.
// The returned map is owned by the workspace and must be treated as read-only.
func (w *Workspace) SourceFilesCached() (map[string][]byte, error) {
	stamps, err := sourceStamps(w.Root, w.sourceStamp)
	if err != nil {
		return nil, err
	}
	if metadataChanged(w.sourceStamp, stamps) {
		if err := w.reloadMetadata(); err != nil {
			return nil, err
		}
		stamps, err = sourceStamps(w.Root, nil)
		if err != nil {
			return nil, err
		}
	}
	return w.buildSourceFiles(stamps)
}

func Fingerprint(files map[string][]byte) string {
	return fingerprint(files, nil)
}

// fingerprint uses the hashes already collected by sourceStamps when they are
// available. The fallback keeps this small public helper useful for callers
// that only have an in-memory source map.
func fingerprint(files map[string][]byte, stamps map[string]fileStamp) string {
	names := make([]string, 0, len(files))
	for n := range files {
		names = append(names, n)
	}
	sort.Strings(names)
	var b strings.Builder
	for _, n := range names {
		b.WriteString(n)
		b.WriteByte(0)
		hash := ""
		if stamps != nil {
			if stamp, ok := stamps[n]; ok {
				hash = stamp.Hash
			}
		}
		if hash == "" {
			hash = office.Hash(files[n])
		}
		b.WriteString(hash)
		b.WriteByte('\n')
	}
	return office.Hash([]byte(b.String()))
}

func sourceStamps(root string, previous map[string]fileStamp) (map[string]fileStamp, error) {
	result := map[string]fileStamp{}
	var sourceBytes int64
	record := func(rel, path string, info os.FileInfo, countSource bool) error {
		if !info.Mode().IsRegular() {
			return fmt.Errorf("workspace source is not a regular file: %s", rel)
		}
		if info.Size() > office.Limit {
			return fmt.Errorf("file budget exceeded: %s", rel)
		}
		if countSource {
			if sourceBytes > office.Limit-info.Size() {
				return fmt.Errorf("workspace source budget exceeded")
			}
			sourceBytes += info.Size()
		}
		changedNS, reliable := fileChangeStamp(path)
		current := fileStamp{Size: info.Size(), ModifiedNS: info.ModTime().UnixNano(), ChangedNS: changedNS}
		if reliable {
			if prior, ok := previous[rel]; ok && prior.Size == current.Size && prior.ModifiedNS == current.ModifiedNS && prior.ChangedNS == current.ChangedNS && prior.Hash != "" {
				current.Hash = prior.Hash
			}
		}
		if current.Hash == "" {
			hash, err := fileHash(path)
			if err != nil {
				return err
			}
			current.Hash = hash
		}
		result[rel] = current
		return nil
	}
	for _, rel := range []string{"project.json", ".wordwright/base.opc", ".wordwright/index.json"} {
		path := filepath.Join(root, filepath.FromSlash(rel))
		info, err := workspaceFileInfo(path, rel)
		if err != nil {
			return nil, err
		}
		if err := record(rel, path, info, false); err != nil {
			return nil, err
		}
	}
	rel := componentLockSource
	path := filepath.Join(root, filepath.FromSlash(rel))
	if info, err := workspaceFileInfo(path, rel); err == nil {
		if hashErr := record(rel, path, info, false); hashErr != nil {
			return nil, hashErr
		}
	} else if !os.IsNotExist(err) {
		return nil, err
	}
	for _, top := range sourceDirectories {
		base := filepath.Join(root, top)
		err := filepath.WalkDir(base, func(path string, entry fs.DirEntry, err error) error {
			if os.IsNotExist(err) {
				return nil
			}
			if err != nil {
				return err
			}
			if entry.Type()&os.ModeSymlink != 0 {
				return fmt.Errorf("source symlink is not accepted")
			}
			if !entry.IsDir() {
				rel, relErr := filepath.Rel(root, path)
				if relErr != nil {
					return relErr
				}
				info, infoErr := entry.Info()
				if infoErr != nil {
					return infoErr
				}
				if hashErr := record(filepath.ToSlash(rel), path, info, true); hashErr != nil {
					return hashErr
				}
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	return result, nil
}

// sourceStamps reads metadata outside Under/Read, so check the final path
// explicitly. Without this guard a metadata symlink could make the resident
// cache hash a file outside the workspace after the workspace was opened.
func workspaceFileInfo(path, rel string) (os.FileInfo, error) {
	info, err := os.Lstat(path)
	if err != nil {
		return nil, err
	}
	if info.Mode()&os.ModeSymlink != 0 {
		return nil, fmt.Errorf("workspace metadata symlink is not accepted: %s", rel)
	}
	if !info.Mode().IsRegular() {
		return nil, fmt.Errorf("workspace metadata is not a regular file: %s", rel)
	}
	return info, nil
}

var workspaceMetadataPaths = [...]string{"project.json", ".wordwright/base.opc", ".wordwright/index.json"}

func metadataChanged(previous, current map[string]fileStamp) bool {
	if previous == nil {
		return false
	}
	for _, path := range workspaceMetadataPaths {
		before, beforeOK := previous[path]
		after, afterOK := current[path]
		if !beforeOK || !afterOK || before.Size != after.Size || before.ModifiedNS != after.ModifiedNS || before.ChangedNS != after.ChangedNS || before.Hash != after.Hash {
			return true
		}
	}
	return false
}

func (w *Workspace) reloadMetadata() error {
	fresh, err := Open(w.Root)
	if err != nil {
		return err
	}
	w.Manifest, w.Index, w.Baseline = fresh.Manifest, fresh.Index, fresh.Baseline
	w.baselineVBA, w.baselineHashes = nil, nil
	w.buildMemo = nil
	w.sourceMemo, w.sourceStamp = nil, nil
	return nil
}

func (w *Workspace) buildSourceFiles(stamps map[string]fileStamp) (map[string][]byte, error) {
	if stamps == nil {
		var err error
		stamps, err = sourceStamps(w.Root, w.sourceStamp)
		if err != nil {
			return nil, err
		}
	}
	if w.sourceMemo == nil {
		files, readErr := w.SourceFiles()
		if readErr != nil {
			return nil, readErr
		}
		if lock, lockErr := Read(w.Root, componentLockSource); lockErr == nil {
			files[componentLockSource] = lock
		} else if !os.IsNotExist(lockErr) {
			return nil, lockErr
		}
		w.sourceMemo, w.sourceStamp = files, stamps
		return files, nil
	}
	files := make(map[string][]byte, len(w.sourceMemo))
	for path, data := range w.sourceMemo {
		files[path] = data
	}
	for path, current := range stamps {
		if strings.HasPrefix(path, ".wordwright/") && path != componentLockSource {
			continue
		}
		if prior, ok := w.sourceStamp[path]; ok && prior == current {
			continue
		}
		data, readErr := Read(w.Root, path)
		if readErr != nil {
			return nil, readErr
		}
		files[path] = data
	}
	for path := range files {
		if _, exists := stamps[path]; !exists {
			delete(files, path)
		}
	}
	w.sourceMemo, w.sourceStamp = files, stamps
	return files, nil
}

func (w *Workspace) baselinePartHashes() map[string]string {
	if w.baselineHashes == nil {
		w.baselineHashes = make(map[string]string, len(w.Baseline.Files))
		for name, data := range w.Baseline.Files {
			w.baselineHashes[name] = office.Hash(data)
		}
	}
	return w.baselineHashes
}

func stamp(path string) (fileStamp, error) {
	info, err := os.Stat(path)
	if err != nil {
		return fileStamp{}, err
	}
	return fileStamp{Size: info.Size(), ModifiedNS: info.ModTime().UnixNano()}, nil
}

func fileHash(path string) (string, error) {
	file, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer file.Close()
	hash := sha256.New()
	if _, err = io.Copy(hash, file); err != nil {
		return "", err
	}
	return hex.EncodeToString(hash.Sum(nil)), nil
}

type ribbonMergeSource struct {
	Component string
	Source    string `json:"source"`
	Target    string `json:"target"`
}

// workspaceRibbonMerges reads only the typed Ribbon composition entries from
// the existing component lock. The lock is provenance, not a second source
// tree; malformed or unsafe entries fail the build before package mutation.
func workspaceRibbonMerges(root string) ([]ribbonMergeSource, error) {
	raw, err := Read(root, ".wordwright/components.json")
	if os.IsNotExist(err) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	var envelope struct {
		Schema     int                        `json:"schema"`
		Components map[string]json.RawMessage `json:"components"`
	}
	if err := json.Unmarshal(raw, &envelope); err != nil {
		return nil, fmt.Errorf("component lock: %w", err)
	}
	if envelope.Schema != 1 || envelope.Components == nil {
		return nil, fmt.Errorf("component lock schema mismatch")
	}
	ids := make([]string, 0, len(envelope.Components))
	for id := range envelope.Components {
		ids = append(ids, id)
	}
	sort.Strings(ids)
	merges := []ribbonMergeSource{}
	for _, id := range ids {
		var component struct {
			RibbonMerges []struct {
				Source string `json:"source"`
				Target string `json:"target"`
			} `json:"ribbon_merges"`
		}
		if err := json.Unmarshal(envelope.Components[id], &component); err != nil {
			return nil, fmt.Errorf("component %s lock: %w", id, err)
		}
		for _, merge := range component.RibbonMerges {
			source := filepath.ToSlash(merge.Source)
			target := filepath.ToSlash(merge.Target)
			if !fs.ValidPath(source) || strings.ContainsAny(source, `\:`) || strings.HasPrefix(strings.ToLower(source), "package/") {
				return nil, fmt.Errorf("component %s Ribbon source path is unsafe: %q", id, merge.Source)
			}
			if !office.SafePart(target) || !strings.HasSuffix(strings.ToLower(target), ".xml") {
				return nil, fmt.Errorf("component %s Ribbon target path is unsafe: %q", id, merge.Target)
			}
			merges = append(merges, ribbonMergeSource{Component: id, Source: source, Target: target})
		}
	}
	return merges, nil
}

func (w *Workspace) Build(output string) (*BuildReport, error) {
	start := time.Now()
	if output == "" {
		output = filepath.Join(w.Root, "dist", w.Manifest.Name+".dotm")
	}
	var e error
	var sourceSnapshot map[string]fileStamp
	output, e = filepath.Abs(output)
	if e != nil {
		return nil, e
	}
	if memo := w.buildMemo; memo != nil && memo.Output == output {
		sources, sourceErr := sourceStamps(w.Root, memo.Sources)
		if sourceErr == nil {
			sourceSnapshot = sources
		}
		artifact, artifactErr := stamp(output)
		evidence, evidenceErr := stamp(filepath.Join(w.Root, "reports", "build.json"))
		if sourceErr == nil && artifactErr == nil && evidenceErr == nil && maps.Equal(sources, memo.Sources) && artifact == memo.Artifact && evidence == memo.Evidence {
			if actualHash, hashErr := fileHash(output); hashErr == nil && actualHash == memo.Report.SHA256 {
				report := memo.Report
				report.Cached = true
				report.DurationMS = float64(time.Since(start).Microseconds()) / 1000
				return &report, nil
			}
		}
		// Something outside this Workspace changed. Reload the immutable
		// package and import index before doing real work so a resident agent
		// observes direct filesystem edits just as a fresh process would.
		immutableChanged := sourceErr != nil
		for _, path := range []string{".wordwright/base.opc", ".wordwright/index.json"} {
			immutableChanged = immutableChanged || sources[path] != memo.Sources[path]
		}
		if immutableChanged {
			fresh, refreshErr := Open(w.Root)
			if refreshErr != nil {
				return nil, refreshErr
			}
			w.Manifest, w.Index, w.Baseline = fresh.Manifest, fresh.Index, fresh.Baseline
			w.baselineVBA, w.baselineHashes, w.sourceMemo, w.sourceStamp = nil, nil, nil, nil
		}
		w.buildMemo = nil
	}
	if sourceSnapshot == nil {
		sourceSnapshot, e = sourceStamps(w.Root, w.sourceStamp)
		if e != nil {
			return nil, e
		}
	}
	if metadataChanged(w.sourceStamp, sourceSnapshot) {
		if e = w.reloadMetadata(); e != nil {
			return nil, e
		}
		sourceSnapshot, e = sourceStamps(w.Root, nil)
		if e != nil {
			return nil, e
		}
	}
	files, e := w.buildSourceFiles(sourceSnapshot)
	if e != nil {
		return nil, e
	}
	// sourceStamps has already paid the hash cost for every source file. Keep
	// that result available to each build decision so a non-cache edit does
	// not hash the same package, VBA, or form bytes again.
	sourceHash := func(name string, data []byte) string {
		if stamp, ok := sourceSnapshot[name]; ok && stamp.Hash != "" {
			return stamp.Hash
		}
		return office.Hash(data)
	}
	var manifest Manifest
	if e = ReadJSON(files["project.json"], &manifest); e != nil {
		return nil, e
	}
	if !office.ValidIdentifier(manifest.Name) {
		return nil, fmt.Errorf("invalid project name")
	}
	w.Manifest = manifest
	finger := fingerprint(files, w.sourceStamp)
	// Cache hits still validate both the source snapshot and artifact bytes.
	if b, e := Read(w.Root, "reports/build.json"); e == nil {
		var prior BuildReport
		if json.Unmarshal(b, &prior) == nil && prior.ToolVersion == Version && ExecutableSHA256() != "" && prior.ToolSHA256 == ExecutableSHA256() && prior.SourceFingerprint == finger && prior.Artifact == output {
			if artifact, e := os.ReadFile(output); e == nil && office.Hash(artifact) == prior.SHA256 {
				prior.Cached = true
				prior.DurationMS = float64(time.Since(start).Microseconds()) / 1000
				return &prior, nil
			}
		}
	}
	original := w.baselinePartHashes()
	// Windows workspaces cannot retain two directory spellings that differ
	// only by case. Map the editable package tree back to the immutable ZIP's
	// exact part names before validation so a relationship such as
	// customXML/item3.xml survives an import/build round trip.
	baselineNames := map[string]string{}
	for name := range w.Baseline.Files {
		folded := strings.ToLower(name)
		if prior, exists := baselineNames[folded]; exists && prior != name {
			return nil, fmt.Errorf("baseline contains case-colliding package parts %s and %s", prior, name)
		}
		baselineNames[folded] = name
	}
	indexNames := map[string]string{}
	for name := range w.Index.Files {
		folded := strings.ToLower(name)
		if prior, exists := indexNames[folded]; exists && prior != name {
			return nil, fmt.Errorf("workspace index contains case-colliding paths %s and %s", prior, name)
		}
		indexNames[folded] = name
	}
	packageFiles := map[string][]byte{}
	packageHashes := map[string]string{}
	packageChanged := false
	for n, b := range files {
		if strings.HasPrefix(n, "package/") {
			part := strings.TrimPrefix(n, "package/")
			if canonical, ok := baselineNames[strings.ToLower(part)]; ok {
				part = canonical
			}
			if prior, exists := packageFiles[part]; exists && !bytes.Equal(prior, b) {
				return nil, fmt.Errorf("package source paths collide after case normalization: %s", part)
			}
			packageFiles[part] = b
			currentHash := sourceHash(n, b)
			packageHashes[part] = currentHash
			expected := w.Index.Files[n]
			if expected == "" {
				if indexed, ok := indexNames[strings.ToLower("package/"+part)]; ok {
					expected = w.Index.Files[indexed]
				}
			}
			if expected != currentHash {
				packageChanged = true
			}
		}
	}
	p := w.Baseline.WithFiles(packageFiles)
	// Keep the list of parts that package composition may rewrite. For every
	// other baseline part, its immutable hash is already authoritative; this
	// avoids hashing unchanged media and custom XML again on a source edit.
	dirtyPackageParts := map[string]bool{}
	// Windows treats package source paths case-insensitively, while the OPC
	// part map is keyed by its original spelling. Resolve Ribbon targets once
	// so a component cannot accidentally create a second part that differs
	// only by case. New targets retain the spelling declared by the first merge.
	ribbonTargetNames := map[string]string{}
	for name := range packageFiles {
		ribbonTargetNames[strings.ToLower(name)] = name
	}
	ribbonMerges, mergeErr := workspaceRibbonMerges(w.Root)
	if mergeErr != nil {
		return nil, mergeErr
	}
	for _, merge := range ribbonMerges {
		fragment, ok := files[merge.Source]
		if !ok {
			return nil, fmt.Errorf("component %s Ribbon source is missing: %s", merge.Component, merge.Source)
		}
		target := merge.Target
		key := strings.ToLower(target)
		if canonical, exists := ribbonTargetNames[key]; exists {
			target = canonical
		} else {
			ribbonTargetNames[key] = target
		}
		if err := p.MergeRibbon(target, fragment); err != nil {
			return nil, fmt.Errorf("component %s Ribbon merge %s -> %s: %w", merge.Component, merge.Source, merge.Target, err)
		}
		dirtyPackageParts[target] = true
		dirtyPackageParts["[Content_Types].xml"] = true
		dirtyPackageParts["_rels/.rels"] = true
	}
	if packageChanged {
		dirtyPackageParts["[Content_Types].xml"] = true
		dirtyPackageParts["_rels/.rels"] = true
		if e = p.ConnectRibbons(); e != nil {
			return nil, e
		}
	}
	mods := []office.Module{}
	formStreams := map[string]map[string][]byte{}
	var v *office.VBA
	cfbOwned := false
	if raw := w.Baseline.Files["word/vbaProject.bin"]; len(raw) > 0 {
		if w.baselineVBA == nil {
			w.baselineVBA, e = office.ReadVBA(raw)
			if e != nil {
				return nil, e
			}
		}
		copyV := *w.baselineVBA
		copyV.Prefix = append([]byte(nil), copyV.Prefix...)
		v = &copyV
	} else {
		v = office.NewVBA(w.Manifest.Name)
	}
	changed := v.Name != w.Manifest.Name || len(v.Modules) == 0 || len(w.Manifest.References) > 0
	v.Name = w.Manifest.Name
	seen := map[string]bool{}
	names := []string{}
	for n := range files {
		if strings.HasPrefix(n, "vba/") {
			names = append(names, n)
		}
	}
	sort.Strings(names)
	for _, n := range names {
		ext := strings.ToLower(filepath.Ext(n))
		if ext != ".bas" && ext != ".cls" && ext != ".vba" {
			return nil, fmt.Errorf("unrecognized source file %s", n)
		}
		name := strings.TrimSuffix(filepath.Base(n), filepath.Ext(n))
		kind, kindErr := w.ModuleKind(name, ext)
		if kindErr != nil {
			return nil, kindErr
		}
		if seen[strings.ToLower(name)] {
			return nil, fmt.Errorf("duplicate component source %s", name)
		}
		seen[strings.ToLower(name)] = true
		source := office.Normalize(string(files[n]))
		currentHash := sourceHash(n, files[n])
		if w.Index.Files[n] != currentHash {
			changed = true
		}
		mod := office.Module{Name: name, Kind: kind, Source: source}
		mods = append(mods, mod)
		if kind == "form" {
			rel := "forms/" + name + ".json"
			raw, ok := files[rel]
			if !ok {
				if _, e := v.CFB.Stream(name + "/f"); e != nil {
					return nil, fmt.Errorf("form %s needs persistent layout source", name)
				}
				continue
			}
			if w.Index.Files[rel] == sourceHash(rel, raw) {
				if _, e := v.CFB.Stream(name + "/f"); e == nil {
					continue
				}
			}
			if !cfbOwned {
				v.CFB = v.CFB.Clone()
				cfbOwned = true
			}
			var design office.Design
			if e = ReadJSON(raw, &design); e != nil {
				return nil, fmt.Errorf("%s: %w", rel, e)
			}
			if design.Name != name {
				return nil, fmt.Errorf("form name mismatch")
			}
			var form *office.Form
			if _, e := v.CFB.Stream(name + "/f"); e == nil {
				form, e = office.ReadForm(v.CFB, name, v.Codepage)
				if e != nil {
					if design.Mode == "preserve" && len(design.Properties) == 0 && len(design.Controls) == 0 && len(design.Remove) == 0 {
						continue
					}
					return nil, e
				}
			} else {
				form, e = office.NewForm(name, v.Codepage)
				if e != nil {
					return nil, e
				}
			}
			if e = form.Apply(design); e != nil {
				return nil, fmt.Errorf("%s: %w", rel, e)
			}
			e = form.WriteBack(v.CFB)
			if e != nil {
				return nil, fmt.Errorf("%s: %w", rel, e)
			}
			changed = true
		}
	}
	if len(mods) != len(v.Modules) {
		changed = true
	}
	if len(mods) == 0 {
		return nil, fmt.Errorf("source workspace must retain its host document module")
	}
	for _, r := range w.Manifest.References {
		if e = v.AddReference(r.Name, r.GUID, r.Version, r.Path, r.Description); e != nil {
			return nil, e
		}
	}
	if v.Name != w.Manifest.Name {
		changed = true
	}
	forms := 0
	for _, m := range mods {
		if m.Kind == "form" {
			forms++
		}
	}
	if changed {
		if b, ok := files["package/word/vbaProject.bin"]; ok && sourceHash("package/word/vbaProject.bin", b) != w.Index.Files["package/word/vbaProject.bin"] {
			return nil, fmt.Errorf("both binary VBA and source changed; choose one authoritative edit path")
		}
		vb, e := v.Rewrite(mods, formStreams)
		if e != nil {
			return nil, e
		}
		if e = p.SetVBA(vb); e != nil {
			return nil, e
		}
		for _, part := range []string{"word/vbaProject.bin", "word/vbaData.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"} {
			dirtyPackageParts[part] = true
		}
	}
	modified, currentHashes := packageChanges(p, original, packageHashes, dirtyPackageParts)
	if len(modified) > 0 && w.Baseline.HasSignatures() {
		if !w.Manifest.DropSignatures {
			return nil, fmt.Errorf("edits invalidate a digital signature; set drop_signatures explicitly or use your signing pipeline")
		}
		if e = dropSignatures(p); e != nil {
			return nil, e
		}
		for n := range p.Files {
			if strings.HasSuffix(n, ".rels") || n == "[Content_Types].xml" {
				dirtyPackageParts[n] = true
			}
		}
		modified, currentHashes = packageChanges(p, original, packageHashes, dirtyPackageParts)
	}
	if e = p.Validate(); e != nil {
		return nil, e
	}
	result, e := p.BytesChangedWithHashes(modified, currentHashes)
	if e != nil {
		return nil, e
	}
	report := &BuildReport{Schema: 2, ToolVersion: Version, ToolSHA256: ExecutableSHA256(), Artifact: output, SHA256: office.Hash(result), SourceFingerprint: finger, Bytes: len(result), NoChange: bytes.Equal(result, w.Baseline.Original), Modules: len(mods), Forms: forms, ModifiedParts: modified, PackageValidated: true}
	if changed {
		report.Warnings = append(report.Warnings, "VBA binary serialization passed source reparse; only the native Word runtime can establish compilation and behavior.")
	}
	if e = AtomicWrite(output, result); e != nil {
		return nil, e
	}
	report.DurationMS = float64(time.Since(start).Microseconds()) / 1000
	if e = Write(w.Root, "reports/build.json", JSON(report), ""); e != nil {
		return nil, e
	}
	if sources, sourceErr := sourceStamps(w.Root, w.sourceStamp); sourceErr == nil {
		if artifact, artifactErr := stamp(output); artifactErr == nil {
			if evidence, evidenceErr := stamp(filepath.Join(w.Root, "reports", "build.json")); evidenceErr == nil {
				w.buildMemo = &buildMemo{Output: output, Sources: sources, Artifact: artifact, Evidence: evidence, Report: *report}
			}
		}
	}
	return report, nil
}

// packageChanges derives the complete changed-part set without rehashing
// immutable baseline parts. packageHashes covers editable package sources
// already hashed during source collection; dirty marks parts rewritten by
// Ribbon/VBA/signature composition and therefore requiring a final hash.
func packageChanges(p *office.Package, original, packageHashes map[string]string, dirty map[string]bool) ([]string, map[string]string) {
	modified := []string{}
	currentHashes := make(map[string]string, len(p.Files))
	for name, data := range p.Files {
		current, known := packageHashes[name]
		if dirty[name] || !known {
			if baseline, exists := original[name]; exists && !dirty[name] && !known {
				current = baseline
			} else {
				current = office.Hash(data)
			}
		}
		currentHashes[name] = current
		if original[name] != current {
			modified = append(modified, name)
		}
	}
	for name := range original {
		if _, ok := p.Files[name]; !ok {
			modified = append(modified, "-"+name)
		}
	}
	sort.Strings(modified)
	return modified, currentHashes
}

func dropSignatures(p *office.Package) error {
	removed := map[string]bool{}
	for n := range p.Files {
		l := strings.ToLower(n)
		if strings.Contains(l, "vbaprojectsignature") || strings.HasPrefix(l, "_xmlsignatures/") {
			delete(p.Files, n)
			removed[n] = true
		}
	}
	for n, b := range p.Files {
		if !strings.HasSuffix(n, ".rels") && n != "[Content_Types].xml" {
			continue
		}
		spans, e := office.XMLSpans(b)
		if e != nil {
			return e
		}
		for i := len(spans) - 1; i >= 0; i-- {
			s := spans[i]
			if s.Depth != 1 {
				continue
			}
			remove := false
			if n == "[Content_Types].xml" {
				remove = removed[strings.TrimPrefix(s.Attribute("", "PartName"), "/")]
			} else {
				typ := strings.ToLower(s.Attribute("", "Type"))
				remove = strings.Contains(typ, "digital-signature") || strings.Contains(typ, "vbaprojectsignature")
			}
			if remove {
				b = append(append([]byte(nil), b[:s.Start]...), b[s.End:]...)
			}
		}
		p.Files[n] = b
	}
	return nil
}

const AgentInstructions = `# WordUp source workspace

Commit editable source, tests, assets and the .wordwright baseline/index; treat dist outputs and reports as generated local evidence. Generated .gitattributes preserves imported bytes and exact XML expectations regardless of Git autocrlf settings while keeping source diffs readable. Do not merge DOTM binaries; merge source and rebuild. Private inputs remain private even when a workspace is Git-ready.
Use the wordup executable. No module imports, VBE typing, or Python setup.

- Edit vba/*.bas, *.cls, and *.vba as ordinary UTF-8 files; module names and VB_Name must agree.
- New .bas files become standard modules; new .cls files become classes; .vba files are UserForm code.
- forms/<name>.json describes persistent native MSForms design, in points. Existing unsupported controls remain opaque.
- package/ is the full original Open XML package, including RibbonX XML, embedded assets and native saved parts. Do not rewrite the ZIP by hand.
- Local component bundles may keep binary assets under assets/ or package/ and declare ribbon_merges in component.json; the normal build composes those fragments into the named customUI part and rejects collisions before copying source.
- Edit package XML directly or use native Word operations for document content, styles, numbering and saved building blocks.
- Use xml.query for exact XML part offsets and hashes, then xml.patch for guarded byte-range edits that preserve every other byte; use xml.verify/xml.compare for direct expected-output checks.
- build performs deterministic package and binary checks; it is not a VBA compiler.
- native execution is local Microsoft Word, never an emulator. Authorize execution only for code the user intends to run. The private desktop is UI separation, NOT a security sandbox.
- Wrap each user-facing editing action in one Application.UndoRecord custom record. For bulk edits, save Application.ScreenUpdating, set it False, and restore the saved value in shared success/error cleanup. Always close an opened undo record; do not blindly restore True when a caller already disabled updates. Read-only actions need no undo record. Test that one undo restores the edited content and that failures restore application state. Capture WordOpenXML outside the editing action and its undo record: a native opening-layout test demonstrated that exporting it during the record disrupted undo grouping.
- Use fresh native acceptance before deployment. Preserve fixture files and assert specific observable behavior; a passing self-test does not establish all other macros work.
- Keep reusable acceptance steps in tests/suite.json and run the test tool with path and fresh=true. Do not build a separate driver script for normal UI workflows.
- For Ribbon actions use ui.invoke with target=document and named:{scope:ribbon,name:caption,role:37 for tabs or 43 for buttons}. For form actions use named:{scope:form,name:caption,window:optional exact form title}. Names are exact; ambiguous selectors fail.
- Follow actions with assertions on document text, formatting, fields or macro diagnostics. A successful action return alone proves no document behavior. Use ui.find with named.wait_ms for bounded control appearance waits. Duplicate names can be narrowed with named.ancestor (exact accessible container name), in addition to window, scope and role; uniqueness remains required. For a modal macro, begin one run with an as task name, handle its owned dialog, then poll that task with eventually_ms and assert both /status equals completed and /error absent. Completed alone is not success; absent requires an existing parent object.
- Keep native tests lean too: use For Each for collection traversal when edits do not invalidate enumeration, and read unchanged Range.Text once before repeated checks. Repeated Paragraphs(i) lookups and full-story reads can dominate the feature under test. For destructive edits, preserve the required reverse order or stable ranges. Measure fixture setup and assertions separately from the editing operation before attributing suite time to the macro.
- xml.snapshot records exact WordOpenXML without saving. Preserve this evidence, but do not assume repeated exports have identical run boundaries: Word pagination can change serialization without a macro. Use explicit behavior assertions; xml.compare is available when exact structure is the intended invariant.
- Failed native suite steps automatically capture owned-window diagnostics and screenshots. Read the errors and images before changing code; do not ask the user to reproduce the failure manually.
- Mac static checks and Windows native tests are not Mac execution evidence.
- Raw XML, native object-model calls and arbitrary VBA are the authoring surface. Unsupported serialization must error rather than substitute an approximation.
- For legal heading/numbering/footnote primitives, consult https://github.com/eliziff/legal-structure-parser. For PDF geometry, reading order and source witnesses, consult https://github.com/eliziff/legal-pdf-parser. These are optional reusable tools, not WordUp runtime dependencies. Prefer native Word evidence, preserve source offsets, and record revisions/licenses for adapted code.
`
