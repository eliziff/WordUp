package verify

import (
	"errors"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"regexp"
	"strings"
)

var procedureDeclaration = regexp.MustCompile(`(?i)^\s*(?:(?:public|private|friend|static)\s+)*(?:sub|function|property\s+(?:get|let|set))\s+([A-Za-z_][A-Za-z0-9_]*)\b`)

func artifactSource(err error, artifact *office.Package) {
	var failure *native.Fault
	if !errors.As(err, &failure) {
		return
	}
	details, ok := failure.Details.(map[string]any)
	if !ok {
		return
	}
	if details["location_kind"] != "native_vbe_selection" {
		details, ok = details["runtime_location"].(map[string]any)
	}
	if !ok || details["location_kind"] != "native_vbe_selection" {
		return
	}
	name, _ := details["module"].(string)
	line, ok := canonical(details["line"]).(float64)
	text, textOK := details["line_text"].(string)
	if name == "" || !ok || !textOK {
		return
	}
	vba, readErr := office.ReadVBA(artifact.Files["word/vbaProject.bin"])
	if readErr != nil {
		details["source_mapping_error"] = readErr.Error()
		return
	}
	for _, module := range vba.Modules {
		if module.Name != name {
			continue
		}
		sourceLine := exportedLine(module.Source, int(line), text)
		if sourceLine == 0 {
			details["source_mapping_error"] = "VBE selection does not match the artifact module; source location was not guessed"
			return
		}
		extension := map[string]string{"standard": ".bas", "class": ".cls", "document": ".cls", "form": ".vba"}[module.Kind]
		details["source_file"] = "vba/" + module.Name + extension
		details["source_line"] = sourceLine
		if procedure := procedureAtLine(module.Source, sourceLine); procedure != "" {
			details["procedure"] = procedure
		}
		details["artifact_module_source_sha256"] = office.Hash([]byte(module.Source))
		return
	}
	details["source_mapping_error"] = "Selected module was not found in the tested artifact"
}

// procedureAtLine identifies the lexical procedure containing a verified
// physical source line. It deliberately does not infer a procedure from a
// declaration in another module or from a malformed signature.
func procedureAtLine(source string, lineNumber int) string {
	if lineNumber <= 0 {
		return ""
	}
	current := ""
	for index, raw := range strings.Split(office.Normalize(source), "\n") {
		line := strings.TrimSpace(strings.SplitN(raw, "'", 2)[0])
		lower := strings.ToLower(line)
		if strings.HasPrefix(lower, "end sub") || strings.HasPrefix(lower, "end function") || strings.HasPrefix(lower, "end property") {
			current = ""
		}
		if match := procedureDeclaration.FindStringSubmatch(line); match != nil {
			current = match[1]
		}
		if index+1 == lineNumber {
			return current
		}
	}
	return ""
}

func exportedLine(source string, visibleLine int, expected string) int {
	count := 0
	for i, line := range strings.Split(office.Normalize(source), "\n") {
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(line)), "attribute ") {
			continue
		}
		count++
		if count == visibleLine {
			if line == strings.TrimSuffix(strings.TrimSuffix(expected, "\n"), "\r") {
				return i + 1
			}
			return 0
		}
	}
	return 0
}
