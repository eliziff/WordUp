// SPDX-License-Identifier: GPL-3.0-or-later
package vbaparse

import (
	"fmt"
	"github.com/antlr4-go/antlr/v4"
	"github.com/eliziff/WordUp/internal/vbaparse/generated"
	"strings"
	"time"
)

type Diagnostic struct {
	Line    int    `json:"line"`
	Column  int    `json:"column"`
	Message string `json:"message"`
}
type errors struct {
	*antlr.DefaultErrorListener
	items []Diagnostic
}

func (e *errors) SyntaxError(_ antlr.Recognizer, _ interface{}, line, column int, msg string, _ antlr.RecognitionException) {
	if len(e.items) < 50 {
		e.items = append(e.items, Diagnostic{line, column + 1, msg})
	}
}

func Parse(source string) (map[string]any, error) {
	return ParseWithConstants(source, nil)
}
func ParseWithConstants(source string, constants map[string]any) (map[string]any, error) {
	if len(source) > 1<<20 {
		return nil, fmt.Errorf("VBA parse source exceeds 1 MiB module budget")
	}
	started := time.Now()
	// The port requires a terminating newline, preserving all original coordinates.
	if !strings.HasSuffix(source, "\n") {
		source += "\n"
	}
	var err error
	source, err = preprocess(source, constants)
	if err != nil {
		return nil, err
	}
	var tree antlr.Tree
	parse := func(mode int) *errors {
		listener := &errors{DefaultErrorListener: antlr.NewDefaultErrorListener(), items: []Diagnostic{}}
		lexer := generated.NewVBALexer(antlr.NewInputStream(source))
		lexer.RemoveErrorListeners()
		lexer.AddErrorListener(listener)
		parser := generated.NewVBAParser(antlr.NewCommonTokenStream(lexer, antlr.TokenDefaultChannel))
		parser.RemoveErrorListeners()
		parser.AddErrorListener(listener)
		parser.BuildParseTrees = true
		parser.GetInterpreter().SetPredictionMode(mode)
		tree = parser.StartRule()
		return listener
	}
	listener := parse(antlr.PredictionModeSLL)
	if len(listener.items) > 0 {
		listener = parse(antlr.PredictionModeLL)
	}
	syntaxValid := len(listener.items) == 0
	if syntaxValid {
		listener.items = append(listener.items, declarationCollisions(tree)...)
	}
	return map[string]any{"syntax_valid": syntaxValid, "declarations_valid": syntaxValid && len(listener.items) == 0, "diagnostics": listener.items, "engine": "Rubberduck VBA grammar / ANTLR 4.13.1 Go", "vba_compiled": false, "word_executed": false, "duration_ms": float64(time.Since(started).Microseconds()) / 1000, "limitations": []string{"Checks module-level constant, variable and procedure collisions; does not resolve project references or types.", "Conditional branches use local-platform defaults or supplied constants. Unsupported conditional intrinsics/operators return errors."}}, nil
}
