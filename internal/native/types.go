// Package native contains actual host integration; it never emulates VBA.
package native

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"runtime"
	"time"
)

type Options struct {
	Visible          bool   `json:"visible,omitempty"`
	Directory        string `json:"directory"`
	Execute          bool   `json:"execute"`
	StartupTimeoutMS int    `json:"startup_timeout_ms,omitempty"`
	MemoryLimitMB    int    `json:"memory_limit_mb,omitempty"`
	CPUPercent       int    `json:"cpu_percent,omitempty"`
}
type Operation struct {
	Op        string         `json:"op"`
	Target    string         `json:"target,omitempty"`
	Member    string         `json:"member,omitempty"`
	Args      []any          `json:"args,omitempty"`
	Named     map[string]any `json:"named,omitempty"`
	Value     any            `json:"value,omitempty"`
	As        string         `json:"as,omitempty"`
	File      string         `json:"file,omitempty"`
	Macro     string         `json:"macro,omitempty"`
	HWND      uint64         `json:"hwnd,omitempty"`
	Child     int            `json:"child,omitempty"`
	Depth     int            `json:"depth,omitempty"`
	Format    string         `json:"format,omitempty"`
	Steps     []Operation    `json:"steps,omitempty"`
	TimeoutMS int            `json:"timeout_ms,omitempty"`
}
type Fault struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Details any    `json:"details,omitempty"`
}

func (e *Fault) Error() string { return e.Code + ": " + e.Message }
func Fail(code, message string, details any) error {
	return &Fault{Code: code, Message: message, Details: details}
}

type Response struct {
	Task       string  `json:"task,omitempty"`
	ID         uint64  `json:"id"`
	Result     any     `json:"result,omitempty"`
	Error      *Fault  `json:"error,omitempty"`
	DurationMS float64 `json:"duration_ms"`
}
type Request struct {
	ID        uint64    `json:"id"`
	Operation Operation `json:"operation"`
}
type Host interface {
	Call(context.Context, Operation) (any, error)
	Close() error
	Info() map[string]any
}

func deadline(ctx context.Context, ms int) (context.Context, context.CancelFunc) {
	if ms <= 0 {
		ms = 120000
	}
	return context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
}
func Describe() map[string]any {
	return map[string]any{"os": runtime.GOOS, "arch": runtime.GOARCH, "engine": "Microsoft Word on the local computer", "native_binary": true, "cloud_required": false, "python_required": false, "word_runtime_bundled": false, "runtime_verification": false, "available": Available()}
}
func fault(e error) *Fault {
	if e == nil {
		return nil
	}
	var f *Fault
	if errors.As(e, &f) {
		if e == f {
			return f
		}
		return &Fault{Code: f.Code, Message: e.Error(), Details: f.Details}
	}
	return &Fault{Code: "native_error", Message: e.Error()}
}
func decodeOperation(b []byte) (Operation, error) {
	var o Operation
	e := json.Unmarshal(b, &o)
	if e == nil && o.Op == "" {
		e = fmt.Errorf("operation required")
	}
	return o, e
}
