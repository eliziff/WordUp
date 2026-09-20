package native

import "fmt"

// Keep actionable dialog text prominent without throwing away the original tree.
func dialogSummary(tree any) map[string]any {
	texts, buttons := []string{}, []string{}
	controls := []any{}
	seen := map[string]bool{}
	var visit func(any)
	visit = func(value any) {
		node, ok := value.(map[string]any)
		if !ok {
			return
		}
		name, _ := node["name"].(string)
		role := fmt.Sprint(node["role"])
		if name != "" && !seen[role+name] {
			if role == "41" {
				texts = append(texts, name)
			}
			if role == "43" {
				buttons = append(buttons, name)
				if node["selector"] != nil {
					controls = append(controls, map[string]any{"name": name, "role": node["role"], "selector": node["selector"]})
				}
			}
			seen[role+name] = true
		}
		children, _ := node["children"].([]any)
		for _, child := range children {
			visit(child)
		}
	}
	visit(tree)
	return map[string]any{"messages": texts, "buttons": buttons, "controls": controls}
}
