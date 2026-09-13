package native

import (
	"fmt"
	"strconv"
	"strings"
)

// Number scratch statements only. User artifacts are never instrumented here.
// Preserve caller-supplied numbers and continuation lines verbatim.
func evaluationSource(module, tag, body string) (string, string) {
	lines := strings.Split(strings.ReplaceAll(body, "\r\n", "\n"), "\n")
	mode := "scratch_body_lines"
	if len(lines) > 65535 {
		mode = "unavailable_line_limit"
	}
	for _, line := range lines {
		s := strings.TrimSpace(line)
		if len(s) > 0 && s[0] >= '0' && s[0] <= '9' {
			mode = "caller_numbered"
			break
		}
	}
	if mode == "scratch_body_lines" {
		continued := false
		for i, line := range lines {
			code := strings.TrimSpace(scratchCode(line))
			// Named labels and directives are left intact. Erl then identifies
			// the last numbered statement, as defined by VBA, not a stack trace.
			label := false
			if colon := strings.IndexByte(code, ':'); colon >= 0 {
				prefix := strings.TrimSpace(code[:colon])
				label = prefix != ""
				for _, c := range prefix {
					if !(c == '_' || c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9') {
						label = false
					}
				}
			}
			// VBA does not allow a line label between Select Case and Case.
			words := strings.Fields(code)
			caseClause := len(words) > 0 && strings.EqualFold(words[0], "Case")
			if !continued && code != "" && !strings.HasPrefix(code, "#") && !label && !caseClause {
				lines[i] = fmt.Sprintf("%d %s", i+1, line)
			}
			// VBA also continues comments ending in space-underscore. Do not
			// turn their following physical line into a numbered statement.
			continued = strings.HasSuffix(code, " _") || code == "_" || strings.HasSuffix(strings.TrimSpace(line), " _")
		}
		body = strings.Join(lines, "\n")
	}
	return "Attribute VB_Name = \"" + module + "\"\nOption Explicit\nPublic Function Evaluate() As Variant\nOn Error GoTo WordUpEvalFailed\n" + body + "\nExit Function\nWordUpEvalFailed:\nEvaluate = Array(\"" + tag + "\", Err.Number, Err.Description, Err.Source, Erl)\nEnd Function\n", mode
}

// Strip comments for continuation detection without interpreting quoted text.
func scratchCode(line string) string {
	quoted, statementStart := false, true
	for i := 0; i < len(line); i++ {
		c := line[i]
		if c == '"' {
			quoted = !quoted
		}
		if quoted {
			continue
		}
		if c == '\'' {
			return line[:i]
		}
		if statementStart && len(line)-i >= 3 && strings.EqualFold(line[i:i+3], "Rem") && (len(line)-i == 3 || line[i+3] == ' ' || line[i+3] == '\t') {
			return line[:i]
		}
		if c == ':' {
			statementStart = true
		} else if c != ' ' && c != '\t' {
			statementStart = false
		}
	}
	return line
}

func evaluationLineDetails(details map[string]any, body, mode string, erl any) {
	details["line_numbering"] = mode
	if mode != "scratch_body_lines" {
		return
	}
	line, err := strconv.Atoi(fmt.Sprint(erl))
	lines := strings.Split(strings.ReplaceAll(body, "\r\n", "\n"), "\n")
	if err == nil && line > 0 && line <= len(lines) {
		details["body_line"] = line
		details["body_line_text"] = lines[line-1]
		details["location_kind"] = "last_executed_numbered_statement"
	}
}
