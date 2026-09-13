//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/eliziff/WordUp/internal/office"
)

type evalEntry struct{ module, source, tag, body, numbering string }
type evalProgram struct {
	entries    map[string]evalEntry
	hash, path string
}

// A verification run loads one separate native add-in for its scratch functions.
// Entries are consumed once, preserving the fresh local/static state of eval.
func (h *wordHost) prepareEvaluation(op Operation) (any, error) {
	if !h.execute {
		return nil, Fail("execution_not_authorized", "Scratch execution requires execute authority", nil)
	}
	if len(op.Steps) == 0 || len(op.Steps) > 1000 || len(h.evalPrograms) >= 16 {
		return nil, fmt.Errorf("scratch program budget exceeded")
	}
	name := "WWEval" + randomID()[:16]
	program := &evalProgram{entries: map[string]evalEntry{}}
	modules := []office.Module{{Name: "ThisDocument", Kind: "document", Source: officeDocumentSource}}
	total := 0
	for i, step := range op.Steps {
		body, ok := step.Value.(string)
		total += len(body)
		if step.Op != "eval" || !ok || total > 1<<20 || step.As == "" {
			return nil, fmt.Errorf("scratch program requires named bounded eval bodies")
		}
		if _, exists := program.entries[step.As]; exists {
			return nil, fmt.Errorf("duplicate scratch entry %s", step.As)
		}
		module := fmt.Sprintf("E%sN%d", name[6:], i)
		tag := randomID()
		source, numbering := evaluationSource(module, tag, body)
		modules = append(modules, office.Module{Name: module, Kind: "standard", Source: source})
		program.entries[step.As] = evalEntry{module, source, tag, body, numbering}
	}
	vb, err := office.NewVBA(name).Rewrite(modules, nil)
	if err != nil {
		return nil, err
	}
	p := office.BlankPackage()
	if err = p.SetVBA(vb); err != nil {
		return nil, err
	}
	b, err := p.Bytes()
	if err != nil {
		return nil, err
	}
	program.path = filepath.Join(h.cfg.Directory, name+".dotm")
	program.hash = office.Hash(b)
	if err = os.WriteFile(program.path, b, 0600); err != nil {
		return nil, err
	}
	if _, err = h.operation(Operation{Op: "addin", File: program.path, As: name}); err != nil {
		return nil, err
	}
	if h.evalPrograms == nil {
		h.evalPrograms = map[string]*evalProgram{}
	}
	h.evalPrograms[name] = program
	return map[string]any{"program": name, "entries": len(modules) - 1, "scratch_sha256": program.hash}, nil
}

func (h *wordHost) executeEvaluation(op Operation) (any, error) {
	if !h.execute {
		return nil, Fail("execution_not_authorized", "Scratch execution requires execute authority", nil)
	}
	p := h.evalPrograms[op.Target]
	if p == nil {
		return nil, fmt.Errorf("unknown scratch program")
	}
	entry, ok := p.entries[op.Member]
	if !ok {
		return nil, fmt.Errorf("unknown or already executed scratch entry")
	}
	delete(p.entries, op.Member)
	value, err := h.app.invoke("Run", 1, []any{entry.module + ".Evaluate"}, nil, h.objects)
	details := map[string]any{"generated_source": entry.source, "scratch_sha256": p.hash, "scratch_path": p.path}
	if err != nil {
		details["cause"] = fault(err)
		return nil, Fail("scratch_vba_failed", err.Error(), details)
	}
	result, err := h.result(&value, op.As)
	if err != nil {
		return nil, err
	}
	array, _ := result.(map[string]any)
	if values, ok := array["array"].([]any); ok && len(values) == 5 && values[0] == entry.tag {
		details["number"] = values[1]
		details["source"] = values[3]
		details["line"] = values[4]
		evaluationLineDetails(details, entry.body, entry.numbering, values[4])
		return nil, Fail("scratch_vba_failed", fmt.Sprint(values[2]), details)
	}
	return map[string]any{"result": result, "executor": "native Microsoft Word VBA", "generated_source": entry.source, "scratch_sha256": p.hash, "scratch_project_separate": true, "scratch_program_reused": true}, nil
}

func (h *wordHost) releaseEvaluation(name string) (any, error) {
	p := h.evalPrograms[name]
	if p == nil {
		return nil, fmt.Errorf("unknown scratch program")
	}
	if err := h.uninstallEvaluation(name, p.path); err != nil {
		return nil, err
	}
	delete(h.evalPrograms, name)
	return true, os.Remove(p.path)
}

func (h *wordHost) uninstallEvaluation(name, path string) error {
	// Loading another scratch add-in can invalidate Word's retained AddIn
	// dispatch. Resolve the exact owned path afresh at the point of removal.
	collection, err := objectProperty(h.app, "AddIns")
	if err != nil {
		return err
	}
	defer collection.release()
	addin, err := objectProperty(collection, "Item", path)
	if err != nil {
		return err
	}
	defer addin.release()
	if err := addin.put("Installed", false); err != nil {
		return err
	}
	if old, ok := h.objects[name]; ok {
		old.release()
		delete(h.objects, name)
	}
	return nil
}
