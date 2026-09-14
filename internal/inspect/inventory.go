package inspect

import (
	"fmt"
	"os"
	pathpkg "path"
	"regexp"
	"sort"
	"strings"

	"github.com/eliziff/WordUp/internal/component"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func declarationParameters(declaration string) []string {
	open, close := strings.IndexByte(declaration, '('), strings.LastIndexByte(declaration, ')')
	if open < 0 || close <= open {
		return nil
	}
	body := declaration[open+1 : close]
	if strings.TrimSpace(body) == "" {
		return []string{}
	}
	parameters := []string{}
	start, depth := 0, 0
	for i, r := range body {
		switch r {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case ',':
			if depth == 0 {
				parameters = append(parameters, strings.TrimSpace(body[start:i]))
				start = i + 1
			}
		}
	}
	return append(parameters, strings.TrimSpace(body[start:]))
}

func parameterType(parameter string) string {
	parameter = strings.TrimSpace(strings.SplitN(parameter, "=", 2)[0])
	for {
		lower := strings.ToLower(parameter)
		removed := false
		for _, prefix := range []string{"optional ", "byval ", "byref ", "paramarray "} {
			if strings.HasPrefix(lower, prefix) {
				parameter = strings.TrimSpace(parameter[len(prefix):])
				removed = true
				break
			}
		}
		if !removed {
			break
		}
	}
	lower := strings.ToLower(parameter)
	if at := strings.LastIndex(lower, " as "); at >= 0 {
		return strings.TrimSpace(parameter[at+4:])
	}
	return ""
}

func ribbonDeclarationMismatch(expected string, actual Symbol) string {
	if !strings.EqualFold(actual.Kind, "Sub") {
		return fmt.Sprintf("declared as %s; Ribbon callbacks must be Public Sub", actual.Kind)
	}
	// Symbols currently records the first declaration line. A continued VBA
	// signature is not enough evidence for an arity warning, so defer it to
	// native compile rather than guessing.
	if strings.Contains(actual.Declaration, "(") && !strings.Contains(actual.Declaration, ")") {
		return ""
	}
	want, got := declarationParameters(expected), declarationParameters(actual.Declaration)
	if len(want) != len(got) {
		return fmt.Sprintf("expects %d parameters but declaration has %d", len(want), len(got))
	}
	for i := range want {
		wantType, gotType := parameterType(want[i]), parameterType(got[i])
		// Some Office callbacks intentionally leave the ByRef return value
		// untyped. Treat that slot as a wildcard while still catching a
		// concrete control/flag/index type mismatch.
		if wantType != "" && gotType != "" && !strings.EqualFold(wantType, gotType) {
			return fmt.Sprintf("parameter %d expects %s but declaration uses %s", i+1, wantType, gotType)
		}
	}
	return ""
}

// These are the event suffixes Word exposes on the standard MSForms controls
// and UserForm itself. The inventory is deliberately lexical: native compile
// and event dispatch remain the authority for a particular Office build.
var formEventSuffixes = map[string]bool{
	"activate": true, "beforeupdate": true, "change": true, "click": true,
	"dblclick": true, "deactivate": true, "enter": true, "exit": true,
	"initialize": true, "keydown": true, "keypress": true, "keyup": true,
	"mousedown": true, "mousemove": true, "mouseup": true, "queryclose": true,
	"remove": true, "resize": true, "scroll": true, "terminate": true,
}

var hotkeyRegistration = regexp.MustCompile(`(?i)\b(?:keybindings\s*\.\s*add|application\s*\.\s*onkey|(^|[^a-z])onkey)\b`)
var contextMenuRegistration = regexp.MustCompile(`(?i)\b(?:commandbars\s*\(|controls\s*\.\s*add|commandbarcontrol)\b`)

type componentLock struct {
	Schema     int `json:"schema"`
	Components map[string]struct {
		ID                 string            `json:"id"`
		Version            string            `json:"version"`
		Files              map[string]string `json:"files"`
		Provenance         string            `json:"provenance"`
		License            string            `json:"license"`
		ManifestSHA256     string            `json:"manifest_sha256"`
		Parameters         map[string]string `json:"parameters"`
		SupportedPlatforms []string          `json:"supported_platforms"`
	} `json:"components"`
}

func publicSymbol(symbol Symbol) bool {
	declaration := strings.ToLower(strings.TrimSpace(symbol.Declaration))
	return !strings.HasPrefix(declaration, "private ") && !strings.HasPrefix(declaration, "friend ")
}

func vbaSourceFiles(files map[string][]byte) map[string]string {
	out := map[string]string{}
	for path := range files {
		if !strings.HasPrefix(path, "vba/") {
			continue
		}
		ext := strings.ToLower(pathpkg.Ext(path))
		if ext != ".bas" && ext != ".cls" && ext != ".vba" {
			continue
		}
		name := strings.TrimSuffix(strings.TrimPrefix(path, "vba/"), ext)
		out[strings.ToLower(name)] = path
	}
	return out
}

func formControlNames(design office.Design, names map[string]bool) {
	var walk func([]office.ControlDesign)
	walk = func(controls []office.ControlDesign) {
		for _, control := range controls {
			if control.Name != "" {
				names[strings.ToLower(control.Name)] = true
			}
			walk(control.Controls)
			walk(control.Pages)
		}
	}
	walk(design.Controls)
	walk(design.Pages)
}

func formEventRow(file string, symbol Symbol, controls map[string]bool, designKnown bool) map[string]any {
	name := symbol.Name
	control, event := "", ""
	if i := strings.LastIndexByte(name, '_'); i > 0 {
		control, event = name[:i], strings.ToLower(name[i+1:])
	}
	lifecycle := strings.EqualFold(name, "UserForm_Initialize") || strings.EqualFold(name, "UserForm_Terminate") ||
		strings.EqualFold(name, "UserForm_QueryClose")
	wiring := "unknown"
	if lifecycle {
		wiring = "form_lifecycle"
	} else if designKnown {
		if controls[strings.ToLower(control)] {
			wiring = "control_present"
		} else {
			wiring = "control_not_declared"
		}
	}
	return map[string]any{
		"file": file, "module": symbol.Module, "name": name, "event": event,
		"control": control, "line": symbol.Line, "declaration": symbol.Declaration,
		"wiring": wiring,
	}
}

func lineRegistrations(files map[string][]byte, pattern *regexp.Regexp, kind string) []map[string]any {
	paths := make([]string, 0)
	for path := range files {
		if strings.HasPrefix(path, "vba/") && (strings.HasSuffix(strings.ToLower(path), ".bas") || strings.HasSuffix(strings.ToLower(path), ".cls") || strings.HasSuffix(strings.ToLower(path), ".vba")) {
			paths = append(paths, path)
		}
	}
	sort.Strings(paths)
	rows := []map[string]any{}
	for _, path := range paths {
		for lineNumber, line := range strings.Split(strings.ReplaceAll(string(files[path]), "\r\n", "\n"), "\n") {
			code := CommentFree(line)
			if !pattern.MatchString(code) {
				continue
			}
			rows = append(rows, map[string]any{"file": path, "line": lineNumber + 1, "kind": kind, "text": strings.TrimSpace(code)})
		}
	}
	return rows
}

func ribbonInventory(files map[string][]byte, publicNames map[string]bool) []map[string]any {
	paths := make([]string, 0)
	for path, data := range files {
		if strings.HasPrefix(path, "package/") && strings.HasSuffix(strings.ToLower(path), ".xml") && strings.Contains(string(data), "customUI") {
			paths = append(paths, path)
		}
	}
	sort.Strings(paths)
	rows := []map[string]any{}
	for _, path := range paths {
		spans, err := office.XMLSpans(files[path])
		if err != nil || len(spans) == 0 || spans[0].Name.Local != "customUI" {
			continue
		}
		if spans[0].Name.Space != "http://schemas.microsoft.com/office/2009/07/customui" && spans[0].Name.Space != "http://schemas.microsoft.com/office/2006/01/customui" {
			continue
		}
		for _, span := range spans {
			if span.Name.Space != spans[0].Name.Space {
				continue
			}
			for _, attr := range span.Attr {
				if attr.Name.Space != "" {
					continue
				}
				declaration, known := office.RibbonCallbackDeclaration(span.Name.Local, attr.Name.Local, attr.Value)
				if !known {
					continue
				}
				name := attr.Value
				if i := strings.LastIndexByte(name, '.'); i >= 0 {
					name = name[i+1:]
				}
				declared := publicNames[strings.ToLower(name)]
				state := "declared"
				if !declared {
					state = "missing"
				}
				rows = append(rows, map[string]any{
					"file": path, "xml_start": span.Start, "control": span.Name.Local,
					"control_id": span.Attribute("", "id"), "attribute": attr.Name.Local,
					"callback": attr.Value, "name": name, "expected_declaration": declaration,
					"state": state,
				})
			}
		}
	}
	return rows
}

func ribbonContextMenus(files map[string][]byte) []map[string]any {
	paths := make([]string, 0)
	for path, data := range files {
		if strings.HasPrefix(path, "package/") && strings.HasSuffix(strings.ToLower(path), ".xml") && strings.Contains(string(data), "customUI") {
			paths = append(paths, path)
		}
	}
	sort.Strings(paths)
	rows := []map[string]any{}
	for _, path := range paths {
		spans, err := office.XMLSpans(files[path])
		if err != nil || len(spans) == 0 || spans[0].Name.Local != "customUI" {
			continue
		}
		for _, span := range spans {
			if span.Name.Space != spans[0].Name.Space || span.Name.Local != "contextMenu" {
				continue
			}
			rows = append(rows, map[string]any{
				"file": path, "xml_start": span.Start, "kind": "Ribbon context menu",
				"id": span.Attribute("", "id"), "idMso": span.Attribute("", "idMso"),
			})
		}
	}
	return rows
}

func installedComponentInventory(root string, diagnostics *[]map[string]any) []map[string]any {
	raw, err := project.Read(root, ".wordwright/components.json")
	if os.IsNotExist(err) {
		return []map[string]any{}
	}
	if err != nil {
		*diagnostics = append(*diagnostics, map[string]any{"severity": "error", "file": ".wordwright/components.json", "message": "cannot read installed component lock: " + err.Error()})
		return []map[string]any{}
	}
	var lock componentLock
	if err := project.ReadJSON(raw, &lock); err != nil {
		*diagnostics = append(*diagnostics, map[string]any{"severity": "error", "file": ".wordwright/components.json", "message": "invalid installed component lock: " + err.Error()})
		return []map[string]any{}
	}
	if lock.Schema != 1 || lock.Components == nil {
		*diagnostics = append(*diagnostics, map[string]any{"severity": "error", "file": ".wordwright/components.json", "message": "installed component lock schema mismatch"})
		return []map[string]any{}
	}
	ids := make([]string, 0, len(lock.Components))
	for id := range lock.Components {
		ids = append(ids, id)
	}
	sort.Strings(ids)
	rows := make([]map[string]any, 0, len(ids))
	for _, id := range ids {
		installed := lock.Components[id]
		state := "clean"
		files := make([]map[string]any, 0, len(installed.Files))
		names := make([]string, 0, len(installed.Files))
		for name := range installed.Files {
			names = append(names, name)
		}
		sort.Strings(names)
		for _, name := range names {
			row := map[string]any{"path": name, "installed_sha256": installed.Files[name]}
			data, readErr := project.Read(root, name)
			switch {
			case os.IsNotExist(readErr):
				row["state"] = "missing"
				state = "modified"
			case readErr != nil:
				row["state"] = "unreadable"
				row["error"] = readErr.Error()
				state = "modified"
			case office.Hash(data) != installed.Files[name]:
				row["state"] = "modified"
				row["current_sha256"] = office.Hash(data)
				state = "modified"
			default:
				row["state"] = "clean"
			}
			files = append(files, row)
		}
		result := map[string]any{"id": id, "version": installed.Version, "state": state, "files": files, "provenance": installed.Provenance, "license": installed.License, "manifest_sha256": installed.ManifestSHA256, "parameters": installed.Parameters, "supported_platforms": installed.SupportedPlatforms, "compatibility": component.Compatibility(installed.SupportedPlatforms)}
		rows = append(rows, result)
		if state != "clean" {
			severity := "warning"
			for _, file := range files {
				if file["state"] == "missing" || file["state"] == "unreadable" {
					severity = "error"
					break
				}
			}
			*diagnostics = append(*diagnostics, map[string]any{"severity": severity, "file": ".wordwright/components.json", "component": id, "message": "installed component source is missing or modified", "state": state})
		}
	}
	return rows
}

func CheckInventory(root string, files map[string][]byte, diagnostics *[]map[string]any) map[string]any {
	paths := vbaSourceFiles(files)
	publicNames := map[string]bool{}
	publicByName := map[string][]Symbol{}
	publicMacros := []map[string]any{}
	formEvents := []map[string]any{}
	for _, path := range paths {
		for _, symbol := range Symbols(strings.TrimSuffix(strings.TrimPrefix(path, "vba/"), pathpkg.Ext(path)), string(files[path])) {
			if publicSymbol(symbol) {
				publicNames[strings.ToLower(symbol.Name)] = true
				publicByName[strings.ToLower(symbol.Name)] = append(publicByName[strings.ToLower(symbol.Name)], symbol)
				if symbol.Kind == "Sub" || symbol.Kind == "Function" {
					publicMacros = append(publicMacros, map[string]any{"file": path, "module": symbol.Module, "name": symbol.Name, "kind": symbol.Kind, "line": symbol.Line, "declaration": symbol.Declaration})
				}
			}
			lower := strings.ToLower(symbol.Name)
			if strings.HasPrefix(lower, "userform_") || (strings.Contains(symbol.Name, "_") && formEventSuffixes[strings.ToLower(symbol.Name[strings.LastIndexByte(symbol.Name, '_')+1:])]) {
				controls := map[string]bool{}
				designKnown := false
				formName := strings.TrimSuffix(strings.TrimPrefix(path, "vba/"), pathpkg.Ext(path))
				if raw, ok := files["forms/"+formName+".json"]; ok {
					var design office.Design
					if err := project.ReadJSON(raw, &design); err == nil {
						designKnown = true
						formControlNames(design, controls)
					}
				}
				if !designKnown && !strings.HasSuffix(strings.ToLower(path), ".vba") && !strings.HasPrefix(lower, "userform_") {
					continue
				}
				formEvents = append(formEvents, formEventRow(path, symbol, controls, designKnown))
				if designKnown && formEvents[len(formEvents)-1]["wiring"] == "control_not_declared" {
					*diagnostics = append(*diagnostics, map[string]any{"severity": "warning", "file": path, "line": symbol.Line, "message": "form event names a control not present in its design", "event": symbol.Name, "control": formEvents[len(formEvents)-1]["control"]})
				}
			}
		}
	}
	sort.Slice(publicMacros, func(i, j int) bool {
		if publicMacros[i]["file"] == publicMacros[j]["file"] {
			return publicMacros[i]["line"].(int) < publicMacros[j]["line"].(int)
		}
		return publicMacros[i]["file"].(string) < publicMacros[j]["file"].(string)
	})
	sort.Slice(formEvents, func(i, j int) bool {
		if formEvents[i]["file"] == formEvents[j]["file"] {
			return formEvents[i]["line"].(int) < formEvents[j]["line"].(int)
		}
		return formEvents[i]["file"].(string) < formEvents[j]["file"].(string)
	})
	for name, symbols := range publicByName {
		if len(symbols) > 1 {
			*diagnostics = append(*diagnostics, map[string]any{"severity": "error", "message": "duplicate public procedure name: " + name, "name": name, "locations": symbols})
		}
	}
	components := installedComponentInventory(root, diagnostics)
	menus := lineRegistrations(files, contextMenuRegistration, "VBA context-menu registration")
	menus = append(menus, ribbonContextMenus(files)...)
	return map[string]any{
		"public_macros":              publicMacros,
		"form_events":                formEvents,
		"ribbon_callbacks":           ribbonInventory(files, publicNames),
		"hotkey_registrations":       lineRegistrations(files, hotkeyRegistration, "VBA hotkey registration"),
		"context_menu_registrations": menus,
		"installed_components":       components,
	}
}
