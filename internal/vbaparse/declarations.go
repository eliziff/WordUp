// SPDX-License-Identifier: GPL-3.0-or-later
package vbaparse

import (
	"fmt"
	"github.com/antlr4-go/antlr/v4"
	"github.com/eliziff/WordUp/internal/vbaparse/generated"
	"strings"
)

// Consume Rubberduck's parsed declaration nodes, respecting inactive branches,
// comments, continuations and procedure scope rather than reparsing source text.
func declarationCollisions(tree antlr.Tree) []Diagnostic {
	type declaration struct {
		kind, name string
		line       int
	}
	seen := map[string][]declaration{}
	out := []Diagnostic{}
	var walk func(antlr.Tree, bool)
	walk = func(node antlr.Tree, moduleDeclaration bool) {
		ctx, ok := node.(antlr.ParserRuleContext)
		if !ok {
			return
		}
		kind := ""
		switch ctx.GetRuleIndex() {
		case generated.VBAParserRULE_moduleConstStmt, generated.VBAParserRULE_moduleVariableStmt:
			moduleDeclaration = true
		case generated.VBAParserRULE_constSubStmt:
			if moduleDeclaration {
				kind = "constant"
			}
		case generated.VBAParserRULE_variableSubStmt:
			if moduleDeclaration {
				kind = "variable"
			}
		case generated.VBAParserRULE_functionStmt:
			kind = "function"
		case generated.VBAParserRULE_subStmt:
			kind = "sub"
		case generated.VBAParserRULE_propertyGetStmt:
			kind = "property get"
		case generated.VBAParserRULE_propertyLetStmt:
			kind = "property let"
		case generated.VBAParserRULE_propertySetStmt:
			kind = "property set"
		}
		if kind != "" {
			for _, child := range ctx.GetChildren() {
				id, ok := child.(antlr.ParserRuleContext)
				if !ok {
					continue
				}
				rule := id.GetRuleIndex()
				if rule != generated.VBAParserRULE_identifier && rule != generated.VBAParserRULE_functionName && rule != generated.VBAParserRULE_subroutineName {
					continue
				}
				name := id.GetText()
				key := strings.ToLower(strings.TrimRight(strings.Trim(name, "[]"), "%&^!#$@"))
				for _, prior := range seen[key] {
					if strings.HasPrefix(kind, "property ") && strings.HasPrefix(prior.kind, "property ") && kind != prior.kind {
						continue
					}
					out = append(out, Diagnostic{id.GetStart().GetLine(), id.GetStart().GetColumn() + 1, fmt.Sprintf("Ambiguous name %q: %s conflicts with %s %q at line %d (VBA names are case-insensitive)", name, kind, prior.kind, prior.name, prior.line)})
					break
				}
				seen[key] = append(seen[key], declaration{kind, name, id.GetStart().GetLine()})
				break
			}
			return // Local declarations and parameters have a separate namespace.
		}
		for _, child := range ctx.GetChildren() {
			walk(child, moduleDeclaration)
		}
	}
	walk(tree, false)
	return out
}
