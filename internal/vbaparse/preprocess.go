// SPDX-License-Identifier: GPL-3.0-or-later
// Port of Rubberduck's conditional compilation visitor: evaluate constants,
// select a branch, and hide inactive tokens while retaining source coordinates.
package vbaparse

import (
	"fmt"
	"github.com/antlr4-go/antlr/v4"
	g "github.com/eliziff/WordUp/internal/vbaparse/conditional"
	"github.com/eliziff/WordUp/internal/vbaparse/generated"
	"math"
	"regexp"
	"runtime"
	"strconv"
	"strings"
)

var directives = regexp.MustCompile(`(?im)^\s*#\s*(if|const)\b`)

func preprocess(source string, supplied map[string]any) (string, error) {
	if !directives.MatchString(source) {
		return source, nil
	}
	constants := map[string]any{"vba7": float64(-1), "win32": float64(-1), "win64": float64(-1), "mac": float64(0)}
	if runtime.GOOS == "darwin" {
		constants["win32"] = float64(0)
		constants["win64"] = float64(0)
		constants["mac"] = float64(-1)
	}
	for k, v := range supplied {
		switch v.(type) {
		case bool, string, float64:
			constants[strings.ToLower(k)] = v
		default:
			return "", fmt.Errorf("invalid compilation constant %s", k)
		}
	}
	listener := &errors{DefaultErrorListener: antlr.NewDefaultErrorListener(), items: []Diagnostic{}}
	lexer := generated.NewVBALexer(antlr.NewInputStream(source))
	lexer.RemoveErrorListeners()
	lexer.AddErrorListener(listener)
	p := g.NewVBAConditionalCompilationParser(antlr.NewCommonTokenStream(lexer, antlr.TokenDefaultChannel))
	p.RemoveErrorListeners()
	p.AddErrorListener(listener)
	tree := p.CompilationUnit()
	if len(listener.items) > 0 {
		return "", fmt.Errorf("conditional compilation line %d column %d: %s", listener.items[0].Line, listener.items[0].Column, listener.items[0].Message)
	}
	input := []rune(source)
	output := make([]rune, len(input))
	for i, c := range input {
		output[i] = ' '
		if c == '\n' || c == '\r' {
			output[i] = c
		}
	}
	var block func(g.ICcBlockContext) error
	block = func(b g.ICcBlockContext) error {
		for _, child := range b.GetChildren() {
			switch c := child.(type) {
			case g.IPhysicalLineContext:
				start, end := c.GetStart().GetStart(), c.GetStop().GetStop()+1
				if start >= 0 && end <= len(input) {
					copy(output[start:end], input[start:end])
				}
			case g.ICcConstContext:
				v, e := ccEval(c.CcExpression(), constants)
				if e != nil {
					return e
				}
				constants[strings.ToLower(c.CcVarLhs().GetText())] = v
			case g.ICcIfBlockContext:
				v, e := ccEval(c.CcIf().CcExpression(), constants)
				if e != nil {
					return e
				}
				taken, e := ccTruth(v)
				if e != nil {
					return e
				}
				if taken {
					if e = block(c.CcBlock()); e != nil {
						return e
					}
				}
				for _, branch := range c.AllCcElseIfBlock() {
					if taken {
						break
					}
					v, e = ccEval(branch.CcElseIf().CcExpression(), constants)
					if e != nil {
						return e
					}
					matches, e := ccTruth(v)
					if e != nil {
						return e
					}
					if matches {
						taken = true
						if e = block(branch.CcBlock()); e != nil {
							return e
						}
					}
				}
				if !taken && c.CcElseBlock() != nil {
					if e = block(c.CcElseBlock().CcBlock()); e != nil {
						return e
					}
				}
			}
		}
		return nil
	}
	if e := block(tree.CcBlock()); e != nil {
		return "", e
	}
	return string(output), nil
}
func ccNumber(v any) (float64, error) {
	switch n := v.(type) {
	case float64:
		return n, nil
	case bool:
		if n {
			return -1, nil
		}
		return 0, nil
	case string:
		return strconv.ParseFloat(n, 64)
	}
	return 0, fmt.Errorf("non-numeric compilation value")
}
func ccTruth(v any) (bool, error) {
	n, e := ccNumber(v)
	if e != nil {
		return false, fmt.Errorf("invalid conditional compilation condition: %w", e)
	}
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return false, fmt.Errorf("non-finite conditional compilation condition")
	}
	return n != 0, nil
}
func ccBool(b bool) float64 {
	if b {
		return -1
	}
	return 0
}
func ccEval(c g.ICcExpressionContext, constants map[string]any) (any, error) {
	if c.Name() != nil {
		name := strings.ToLower(strings.Trim(c.Name().GetText(), "[]"))
		if v, ok := constants[name]; ok {
			return v, nil
		}
		return float64(0), nil
	}
	if c.Literal() != nil {
		s := c.Literal().GetText()
		switch strings.ToLower(s) {
		case "true":
			return float64(-1), nil
		case "false", "empty":
			return float64(0), nil
		}
		if strings.HasPrefix(s, "\"") {
			return strings.ReplaceAll(s[1:len(s)-1], "\"\"", "\""), nil
		}
		if strings.HasPrefix(strings.ToLower(s), "&h") {
			n, e := strconv.ParseInt(strings.TrimRight(s[2:], "&"), 16, 64)
			return float64(n), e
		}
		if strings.HasPrefix(strings.ToLower(s), "&o") {
			n, e := strconv.ParseInt(strings.TrimRight(s[2:], "&"), 8, 64)
			return float64(n), e
		}
		return strconv.ParseFloat(strings.TrimRight(s, "%&!#@"), 64)
	}
	if f := c.IntrinsicFunction(); f != nil {
		v, e := ccEval(f.CcExpression(), constants)
		if e != nil {
			return nil, e
		}
		name := strings.ToLower(f.IntrinsicFunctionName().GetText())
		if name == "cstr" {
			return fmt.Sprint(v), nil
		}
		if name == "len" {
			return float64(len([]rune(fmt.Sprint(v)))), nil
		}
		n, e := ccNumber(v)
		if e != nil {
			return nil, e
		}
		switch name {
		case "abs":
			return math.Abs(n), nil
		case "int":
			return math.Floor(n), nil
		case "fix":
			return math.Trunc(n), nil
		case "sgn":
			if n == 0 {
				return float64(0), nil
			}
			return math.Copysign(1, n), nil
		case "cbool":
			return ccBool(n != 0), nil
		case "cint", "clng", "clnglng", "clngptr":
			return math.RoundToEven(n), nil
		case "cdbl", "csng", "cvar":
			return n, nil
		}
		return nil, fmt.Errorf("unsupported conditional intrinsic %s", name)
	}
	children := c.AllCcExpression()
	if len(children) == 0 {
		return nil, fmt.Errorf("unsupported conditional expression %s", c.GetText())
	}
	a, e := ccEval(children[0], constants)
	if e != nil {
		return nil, e
	}
	op := ""
	for _, child := range c.GetChildren() {
		if t, ok := child.(antlr.TerminalNode); ok {
			text := strings.ToLower(t.GetText())
			if text != "(" && text != ")" {
				op = text
				break
			}
		}
	}
	if len(children) == 1 {
		if op == "" {
			return a, nil
		}
		n, e := ccNumber(a)
		if e != nil {
			return nil, e
		}
		switch op {
		case "-":
			return -n, nil
		case "not":
			return float64(^int64(n)), nil
		}
	}
	if len(children) != 2 {
		return nil, fmt.Errorf("unsupported conditional operator %s", op)
	}
	b, e := ccEval(children[1], constants)
	if e != nil {
		return nil, e
	}
	if op == "&" {
		return fmt.Sprint(a) + fmt.Sprint(b), nil
	}
	if sa, ok := a.(string); ok {
		if sb, ok := b.(string); ok {
			switch op {
			case "+":
				return sa + sb, nil
			case "=":
				return ccBool(sa == sb), nil
			case "<>":
				return ccBool(sa != sb), nil
			case "<":
				return ccBool(sa < sb), nil
			case ">":
				return ccBool(sa > sb), nil
			case "<=":
				return ccBool(sa <= sb), nil
			case ">=":
				return ccBool(sa >= sb), nil
			}
		}
	}
	x, e := ccNumber(a)
	if e != nil {
		return nil, e
	}
	y, e := ccNumber(b)
	if e != nil {
		return nil, e
	}
	switch op {
	case "+":
		return x + y, nil
	case "-":
		return x - y, nil
	case "*":
		return x * y, nil
	case "/":
		if y != 0 {
			return x / y, nil
		}
	case "\\":
		if math.RoundToEven(y) != 0 {
			return math.Trunc(math.RoundToEven(x) / math.RoundToEven(y)), nil
		}
	case "mod":
		if y != 0 {
			return math.Mod(math.RoundToEven(x), math.RoundToEven(y)), nil
		}
	case "^":
		return math.Pow(x, y), nil
	case "=":
		return ccBool(x == y), nil
	case "<>":
		return ccBool(x != y), nil
	case "<":
		return ccBool(x < y), nil
	case ">":
		return ccBool(x > y), nil
	case "<=":
		return ccBool(x <= y), nil
	case ">=":
		return ccBool(x >= y), nil
	case "and":
		return float64(int64(x) & int64(y)), nil
	case "or":
		return float64(int64(x) | int64(y)), nil
	case "xor":
		return float64(int64(x) ^ int64(y)), nil
	case "eqv":
		return float64(^(int64(x) ^ int64(y))), nil
	case "imp":
		return float64(^int64(x) | int64(y)), nil
	}
	return nil, fmt.Errorf("invalid or unsupported conditional operation %s", c.GetText())
}
