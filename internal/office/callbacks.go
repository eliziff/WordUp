// Callback mapping adapted from Office RibbonX Editor's CallbacksBuilder.cs.
// Copyright (c) 2019-2020 Fernando Andreu. MIT; see RIBBONX-LICENSE.
package office

import (
	"fmt"
	"regexp"
	"strings"
)

var callbackIdentifier = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)

// RibbonCallbackDeclaration supplies the upstream callback shape, not a VBA
// type-check verdict. Qualified bindings still declare an unqualified VBA name.
func RibbonCallbackDeclaration(control, attribute, binding string) (string, bool) {
	args, known := callbackArguments(control, attribute)
	if i := strings.LastIndex(binding, "."); i >= 0 {
		binding = binding[i+1:]
	}
	if !known || !callbackIdentifier.MatchString(binding) {
		return "", false
	}
	return "Public Sub " + binding + "(" + args + ")", true
}

func RibbonCallbacks(data []byte) (map[string]any, error) {
	nodes, err := XMLSpans(data)
	if err != nil {
		return nil, err
	}
	signatures := map[string]string{}
	callbacks := []map[string]string{}
	var source strings.Builder
	for _, node := range nodes {
		if node.Name.Space != "http://schemas.microsoft.com/office/2009/07/customui" && node.Name.Space != "http://schemas.microsoft.com/office/2006/01/customui" {
			continue
		}
		for _, attr := range node.Attr {
			if attr.Name.Space != "" {
				continue
			}
			args, known := callbackArguments(node.Name.Local, attr.Name.Local)
			if !known {
				if strings.HasPrefix(attr.Name.Local, "get") || strings.HasPrefix(attr.Name.Local, "on") {
					return nil, fmt.Errorf("unsupported Ribbon callback %s.%s", node.Name.Local, attr.Name.Local)
				}
				continue
			}
			name := attr.Value
			if i := strings.LastIndex(name, "."); i >= 0 {
				name = name[i+1:]
			}
			if !callbackIdentifier.MatchString(name) {
				return nil, fmt.Errorf("invalid VBA callback identifier %q", name)
			}
			key := strings.ToLower(name)
			if old, exists := signatures[key]; exists {
				if old != args {
					return nil, fmt.Errorf("callback %s is used with incompatible signatures: (%s) and (%s)", name, old, args)
				}
				continue
			}
			signatures[key] = args
			signature, _ := RibbonCallbackDeclaration(node.Name.Local, attr.Name.Local, name)
			callbacks = append(callbacks, map[string]string{"name": name, "signature": signature, "control": node.Name.Local, "attribute": attr.Name.Local})
			source.WriteString(signature + "\nEnd Sub\n\n")
		}
	}
	return map[string]any{"source": source.String(), "callbacks": callbacks, "engine": "Office RibbonX Editor callback mapping", "vba_compiled": false}, nil
}
func callbackArguments(control, attribute string) (string, bool) {
	const c = "control As IRibbonControl"
	switch attribute {
	case "onLoad":
		return "ribbon As IRibbonUI", true
	case "onShow", "onHide":
		return "contextObject As Object", true
	case "loadImage":
		return "imageID As String, ByRef returnedVal", true
	case "onChange":
		return c + ", text As String", true
	case "onAction":
		switch control {
		case "dropDown", "gallery":
			return c + ", id As String, index As Integer", true
		case "command":
			return c + ", ByRef cancelDefault", true
		case "toggleButton", "checkBox":
			return c + ", pressed As Boolean", true
		}
		return c, true
	case "getItemLabel", "getItemTooltip", "getItemID", "getItemImage":
		return c + ", index As Integer, ByRef returnedVal", true
	case "getEnabled", "getVisible", "getPressed", "getShowLabel", "getShowImage", "getLabel", "getScreentip", "getSupertip", "getDescription", "getKeytip", "getSelectedItemID", "getImageMso", "getContent", "getText", "getTitle", "getTarget", "getHelperText", "getImage", "getItemCount", "getItemIndex", "getItemHeight", "getItemWidth", "getSelectedItemIndex", "getSize", "getItemSize", "getStyle":
		return c + ", ByRef returnedVal", true
	}
	return "", false
}
