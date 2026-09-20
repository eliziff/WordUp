//go:build windows && (amd64 || arm64)

package native

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf16"
	"unsafe"
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
	if op.File != "" {
		if !h.execute {
			return nil, Fail("execution_not_authorized", "Writing XML evidence requires execute authority", nil)
		}
		if value.VT != 8 {
			return nil, fmt.Errorf("WordOpenXML did not return XML")
		}
		// The response limit protects inline results, not file-backed evidence.
		// Keep Word's BSTR alive while decoding bounded UTF-16 chunks directly
		// to disk; do not materialize another full document string or span tree.
		p := uintptr(value.Value)
		n, _, _ := sysLen.Call(p)
		return exportXMLFile(&xmlBSTRReader{remaining: unsafe.Slice((*uint16)(unsafe.Pointer(p)), int(n))}, op.File)
	}
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
	report["xml"] = xml
	return report, nil
}

type xmlBSTRReader struct {
	remaining []uint16
	chunk     *strings.Reader
}

func (r *xmlBSTRReader) Read(p []byte) (int, error) {
	if len(p) == 0 {
		return 0, nil
	}
	if r.chunk == nil || r.chunk.Len() == 0 {
		if len(r.remaining) == 0 {
			return 0, io.EOF
		}
		n := min(len(r.remaining), 32*1024)
		if n < len(r.remaining) && r.remaining[n-1] >= 0xd800 && r.remaining[n-1] <= 0xdbff {
			n-- // Never split a surrogate pair across conversion chunks.
		}
		r.chunk = strings.NewReader(string(utf16.Decode(r.remaining[:n])))
		r.remaining = r.remaining[n:]
	}
	return r.chunk.Read(p)
}

func exportXMLFile(source io.Reader, path string) (any, error) {
	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		return nil, err
	}
	f, err := os.CreateTemp(filepath.Dir(path), ".wordup-xml-*")
	if err != nil {
		return nil, err
	}
	defer os.Remove(f.Name())
	defer f.Close()
	hash := sha256.New()
	decoder := xml.NewDecoder(io.TeeReader(source, io.MultiWriter(f, hash)))
	parts := []string{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		if start, ok := token.(xml.StartElement); ok && start.Name == (xml.Name{Space: "http://schemas.microsoft.com/office/2006/xmlPackage", Local: "part"}) {
			for _, attr := range start.Attr {
				if attr.Name == (xml.Name{Space: start.Name.Space, Local: "name"}) {
					parts = append(parts, attr.Value)
				}
			}
		}
	}
	info, err := f.Stat()
	if err != nil {
		return nil, err
	}
	if err := f.Close(); err != nil {
		return nil, err
	}
	if err := os.Rename(f.Name(), path); err != nil {
		return nil, err
	}
	return map[string]any{"file": path, "sha256": hex.EncodeToString(hash.Sum(nil)), "bytes": info.Size(), "parts": parts, "source": "Microsoft Word WordOpenXML", "normalized": false, "document_saved": false}, nil
}
