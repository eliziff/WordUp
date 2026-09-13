//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
)

// Export Word's live Flat OPC without saving, normalizing or rewriting it.
func (h *wordHost) xmlSnapshot(op Operation) (any, error) {
	if op.Named["failure_capture"] == true {
		record, err := objectProperty(h.app, "UndoRecord")
		if err != nil {
			return nil, err
		}
		defer record.release()
		level, err := record.get("CustomRecordLevel")
		if err != nil {
			return nil, err
		}
		defer level.clear()
		n, err := level.value(0)
		if err != nil {
			return nil, err
		}
		if fmt.Sprint(n) != "0" {
			return nil, Fail("xml_capture_undo_active", "XML capture skipped to preserve the open undo record", map[string]any{"custom_record_level": n})
		}
		doc, err := objectProperty(h.app, "ActiveDocument")
		if err != nil {
			return nil, err
		}
		defer doc.release()
		return h.exportXML(doc, op)
	}
	if op.Target == "" || op.Target == "app" {
		return nil, fmt.Errorf("xml.snapshot requires an explicit document or range target")
	}
	object, err := h.object(op.Target)
	if err != nil {
		return nil, err
	}
	return h.exportXML(object, op)
}

func (h *wordHost) exportXML(object dispatch, op Operation) (any, error) {
	value, err := object.get("WordOpenXML")
	if err != nil {
		// A Content range is not an equivalent document export: it can omit
		// document properties and custom XML. Preserve the actual failure.
		return nil, err
	}
	defer value.clear()
	result, err := value.value(0)
	if err != nil {
		return nil, err
	}
	xml, ok := result.(string)
	if !ok {
		return nil, fmt.Errorf("WordOpenXML did not return XML")
	}
	raw := []byte(xml)
	spans, err := office.XMLSpans(raw)
	if err != nil {
		return nil, err
	}
	parts := []string{}
	for _, span := range spans {
		if span.Name.Local == "part" && span.Name.Space == "http://schemas.microsoft.com/office/2006/xmlPackage" {
			parts = append(parts, span.Attribute("http://schemas.microsoft.com/office/2006/xmlPackage", "name"))
		}
	}
	report := map[string]any{"sha256": office.Hash(raw), "bytes": len(raw), "parts": parts, "source": "Microsoft Word WordOpenXML", "normalized": false, "document_saved": false}
	if op.File == "" {
		report["xml"] = xml
		return report, nil
	}
	if !h.execute {
		return nil, Fail("execution_not_authorized", "Writing XML evidence requires execute authority", nil)
	}
	if err = os.MkdirAll(filepath.Dir(op.File), 0700); err != nil {
		return nil, err
	}
	if err = os.WriteFile(op.File, raw, 0600); err != nil {
		return nil, err
	}
	report["file"] = op.File
	return report, nil
}
