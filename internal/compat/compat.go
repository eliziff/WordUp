// Package compat is a conservative, branch-aware portability review, never a
// replacement for the native VBA compiler. Unknown expressions keep both arms.
package compat

import (
	"github.com/eliziff/WordUp/internal/inspect"
	"github.com/eliziff/WordUp/internal/project"
	"regexp"
	"strings"
)

type truth int

const (
	no      truth = 0
	yes     truth = 1
	unknown truth = 2
)

func not(a truth) truth {
	if a == unknown {
		return unknown
	}
	if a == yes {
		return no
	}
	return yes
}
func and(a, b truth) truth {
	if a == no || b == no {
		return no
	}
	if a == unknown || b == unknown {
		return unknown
	}
	return yes
}
func or(a, b truth) truth {
	if a == yes || b == yes {
		return yes
	}
	if a == unknown || b == unknown {
		return unknown
	}
	return no
}

type expression struct {
	tokens []string
	i      int
	env    map[string]truth
}

var lex = regexp.MustCompile(`(?i)[a-z_][a-z_0-9]*|-?\d+|<>|=|\(|\)|\S`)

func evaluate(s string, env map[string]truth) truth {
	p := expression{tokens: lex.FindAllString(strings.ToLower(strings.TrimSpace(s)), -1), env: env}
	v := p.expr()
	if p.i != len(p.tokens) {
		return unknown
	}
	return v
}
func (p *expression) take(s string) bool {
	if p.i < len(p.tokens) && p.tokens[p.i] == s {
		p.i++
		return true
	}
	return false
}
func (p *expression) expr() truth {
	a := p.term()
	for p.take("or") {
		a = or(a, p.term())
	}
	return a
}
func (p *expression) term() truth {
	a := p.factor()
	for p.take("and") {
		a = and(a, p.factor())
	}
	return a
}
func (p *expression) factor() truth {
	if p.take("not") {
		return not(p.factor())
	}
	if p.take("(") {
		a := p.expr()
		if !p.take(")") {
			return unknown
		}
		return a
	}
	if p.i >= len(p.tokens) {
		return unknown
	}
	s := p.tokens[p.i]
	p.i++
	var a truth
	switch s {
	case "true", "-1", "1":
		a = yes
	case "false", "0":
		a = no
	default:
		var ok bool
		a, ok = p.env[s]
		if !ok {
			a = unknown
		}
	}
	if p.take("=") {
		b := p.factor()
		if a == unknown || b == unknown {
			return unknown
		}
		if a == b {
			return yes
		}
		return no
	}
	if p.take("<>") {
		b := p.factor()
		if a == unknown || b == unknown {
			return unknown
		}
		if a != b {
			return yes
		}
		return no
	}
	return a
}

type Diagnostic struct {
	File                 string `json:"file"`
	Line                 int    `json:"line"`
	Target               string `json:"target"`
	Severity             string `json:"severity"`
	Rule                 string `json:"rule"`
	Message              string `json:"message"`
	ConditionalCertainty string `json:"conditional_certainty"`
}
type Profile struct {
	Name      string          `json:"name"`
	Mac       bool            `json:"mac"`
	Bits      int             `json:"bits"`
	Constants map[string]bool `json:"constants,omitempty"`
}

var profiles = []Profile{{Name: "windows-vba7-64", Bits: 64, Constants: map[string]bool{"win32": true, "win64": true, "mac": false, "vba7": true, "vba6": true}}, {Name: "windows-vba7-32", Bits: 32, Constants: map[string]bool{"win32": true, "win64": false, "mac": false, "vba7": true, "vba6": true}}, {Name: "mac-vba7-64", Mac: true, Bits: 64, Constants: map[string]bool{"mac": true, "vba7": true, "vba6": true}}}
var ptrsafe = regexp.MustCompile(`(?i)\bPtrSafe\b`)
var asLong = regexp.MustCompile(`(?i)\bAs\s+Long\b`)
var declare = regexp.MustCompile(`(?i)\bDeclare\b.*\b(?:Sub|Function)\b`)
var winlib = regexp.MustCompile(`(?i)\bLib\s+"[^"]*(?:\.dll|kernel32|user32|gdi32|ole32|oleaut32|advapi32|shell32|comdlg32|ntdll)[^"]*"`)
var com = regexp.MustCompile(`(?i)\b(?:CreateObject|GetObject)\s*\(`)
var constDecl = regexp.MustCompile(`(?i)^#const\s+(\w+)\s*=\s*(.*)$`)

type frame struct{ parent, seen truth }

func Source(file, src string, pr Profile) []Diagnostic {
	env := map[string]truth{"win16": no}
	for k, v := range pr.Constants {
		if v {
			env[strings.ToLower(k)] = yes
		} else {
			env[strings.ToLower(k)] = no
		}
	}
	active := yes
	stack := []frame{}
	out := []Diagnostic{}
	lines := strings.Split(strings.ReplaceAll(src, "\r\n", "\n"), "\n")
	emit := func(i int, severity, rule, msg string) {
		certainty := "active"
		if active == unknown {
			certainty = "possible; condition unresolved"
		}
		out = append(out, Diagnostic{file, i + 1, pr.Name, severity, rule, msg, certainty})
	}
	for i := 0; i < len(lines); i++ {
		start := i
		l := strings.TrimSpace(inspect.CommentFree(lines[i]))
		for strings.HasSuffix(l, " _") && i+1 < len(lines) {
			i++
			l = strings.TrimSuffix(l, " _") + " " + strings.TrimSpace(inspect.CommentFree(lines[i]))
		}
		low := strings.ToLower(l)
		if m := constDecl.FindStringSubmatch(l); m != nil {
			if active == yes {
				env[strings.ToLower(m[1])] = evaluate(m[2], env)
			} else if active == unknown {
				env[strings.ToLower(m[1])] = unknown
			}
			continue
		}
		if strings.HasPrefix(low, "#if ") {
			cond := evaluate(strings.TrimSuffix(strings.TrimSpace(low[4:]), "then"), env)
			stack = append(stack, frame{active, cond})
			active = and(active, cond)
			if cond == unknown {
				emit(start, "warning", "conditional-unresolved", "Unknown conditional expression; both branches remain under review")
			}
			continue
		}
		if strings.HasPrefix(low, "#elseif ") {
			if len(stack) == 0 {
				emit(start, "error", "conditional-structure", "#ElseIf without #If")
				continue
			}
			fr := &stack[len(stack)-1]
			cond := evaluate(strings.TrimSuffix(strings.TrimSpace(low[8:]), "then"), env)
			active = and(fr.parent, and(not(fr.seen), cond))
			fr.seen = or(fr.seen, cond)
			continue
		}
		if low == "#else" {
			if len(stack) == 0 {
				emit(start, "error", "conditional-structure", "#Else without #If")
				continue
			}
			fr := &stack[len(stack)-1]
			active = and(fr.parent, not(fr.seen))
			fr.seen = yes
			continue
		}
		if low == "#end if" || low == "#endif" {
			if len(stack) == 0 {
				emit(start, "error", "conditional-structure", "#End If without #If")
				continue
			}
			active = stack[len(stack)-1].parent
			stack = stack[:len(stack)-1]
			continue
		}
		if active == no {
			continue
		}
		if declare.MatchString(l) {
			if pr.Bits == 64 && !ptrsafe.MatchString(l) {
				emit(start, "error", "declare-ptrsafe", "64-bit VBA declarations require PtrSafe; audit pointer/handle types separately")
			}
			if pr.Mac && winlib.MatchString(l) {
				emit(start, "error", "windows-library-on-mac", "Windows DLL import is reachable in the Mac review profile; guard with #If Not Mac and provide a native Mac implementation")
			}
			if pr.Bits == 64 && asLong.MatchString(l) {
				emit(start, "review", "native-abi-width", "Review each Long in a native declaration: DWORD/int32 may be correct; pointers/handles generally need LongPtr. No automatic replacement was made")
			}
		}
		if pr.Mac && com.MatchString(l) {
			emit(start, "review", "automation-dependency", "External Automation/COM dependency needs an actual Mac implementation or native dependency test")
		}
		if pr.Mac && strings.Contains(l, `:\`) {
			emit(start, "review", "windows-path", "Windows-style path literal; use platform-aware paths and test sandbox file access")
		}
		if !pr.Mac && (strings.Contains(low, "applescripttask") || strings.Contains(low, "macscript(")) {
			emit(start, "error", "mac-call-on-windows", "Mac-specific scripting call is reachable in a Windows profile")
		}
	}
	if len(stack) != 0 {
		emit(len(lines)-1, "error", "conditional-structure", "Unclosed conditional block")
	}
	return out
}
func Review(w *project.Workspace) (map[string]any, error) {
	files, e := w.SourceFiles()
	if e != nil {
		return nil, e
	}
	out := []Diagnostic{}
	for n, b := range files {
		if !strings.HasPrefix(n, "vba/") {
			continue
		}
		for _, p := range profiles {
			out = append(out, Source(n, string(b), p)...)
		}
	}
	return map[string]any{"profiles": profiles, "diagnostics": out, "mac_runtime_verified": false, "vba_compiled": false, "scope": "Conservative lexical/conditional portability review. Mac Win32/Win64 flags remain unknown until observed on the target host. Unknown expressions do not suppress branches; arbitrary VBA semantics are not interpreted."}, nil
}
func ProbeSource() string {
	return `Attribute VB_Name = "PlatformProbe"
Option Explicit
Public Function PlatformReport() As String
    Dim s As String
#If Mac Then
    s = "Mac=1;"
#Else
    s = "Mac=0;"
#End If
#If Win32 Then
    s = s & "Win32=1;"
#Else
    s = s & "Win32=0;"
#End If
#If Win64 Then
    s = s & "Win64=1;"
#Else
    s = s & "Win64=0;"
#End If
#If VBA7 Then
    s = s & "VBA7=1;"
#Else
    s = s & "VBA7=0;"
#End If
    PlatformReport = s & "Version=" & Application.Version
End Function
`
}
