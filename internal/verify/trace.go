package verify

import (
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"io"
	"path/filepath"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

// SerialTrace is evidence emitted by opt-in instrumentation, not an automatic
// interception of arbitrary VBA. Coverage describes observed paths, not proof
// that uninstrumented paths do not exist. Data preserves exact project values.
type SerialTrace struct {
	Schema   int          `json:"schema"`
	Complete bool         `json:"complete"`
	Coverage []string     `json:"coverage"`
	Events   []TraceEvent `json:"events"`
}

type TraceEvent struct {
	Sequence  int            `json:"sequence"`
	Kind      string         `json:"kind"`
	Stage     string         `json:"stage"`
	Document  string         `json:"document"`
	Story     string         `json:"story"`
	Operation string         `json:"operation,omitempty"`
	Data      map[string]any `json:"data"`
}

type TraceEvidence struct {
	File     string   `json:"file"`
	SHA256   string   `json:"sha256"`
	Events   int      `json:"events"`
	Coverage []string `json:"coverage"`
}

func decodeTrace(raw []byte) (*SerialTrace, error) {
	if len(raw) > 64<<20 {
		return nil, fmt.Errorf("serial trace exceeds 64 MiB")
	}
	// FSO's Unicode export is UTF-16LE. Preserve raw hashes while decoding it.
	if bytes.HasPrefix(raw, []byte{255, 254}) {
		if len(raw)%2 != 0 {
			return nil, fmt.Errorf("truncated UTF-16 trace")
		}
		units := make([]uint16, (len(raw)-2)/2)
		for i := range units {
			units[i] = binary.LittleEndian.Uint16(raw[2+i*2:])
		}
		for i := 0; i < len(units); i++ {
			if units[i] >= 0xD800 && units[i] <= 0xDBFF {
				if i+1 == len(units) || units[i+1] < 0xDC00 || units[i+1] > 0xDFFF {
					return nil, fmt.Errorf("invalid UTF-16 trace")
				}
				i++
			} else if units[i] >= 0xDC00 && units[i] <= 0xDFFF {
				return nil, fmt.Errorf("invalid UTF-16 trace")
			}
		}
		raw = []byte(string(utf16.Decode(units)))
	}
	raw = bytes.TrimPrefix(raw, []byte{239, 187, 191})
	if !utf8.Valid(raw) {
		return nil, fmt.Errorf("invalid UTF-8 trace")
	}
	var trace SerialTrace
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.DisallowUnknownFields()
	decoder.UseNumber()
	if err := decoder.Decode(&trace); err != nil {
		return nil, err
	}
	if err := decoder.Decode(new(any)); err != io.EOF {
		return nil, fmt.Errorf("trailing trace data")
	}
	if trace.Schema != 1 || !trace.Complete || len(trace.Coverage) == 0 || trace.Events == nil {
		return nil, fmt.Errorf("trace needs schema 1, complete=true, coverage and events")
	}
	for _, path := range trace.Coverage {
		if path == "" {
			return nil, fmt.Errorf("empty trace coverage")
		}
	}
	pending := map[string]TraceEvent{}
	used := map[string]bool{}
	for i, event := range trace.Events {
		if event.Sequence != i+1 || event.Stage == "" || event.Document == "" || event.Story == "" || event.Data == nil {
			return nil, fmt.Errorf("invalid trace event %d: sequence, stage, document, story and data required", i+1)
		}
		coordinates := map[string]int64{}
		for _, key := range []string{"start", "end", "story_type", "story_length"} {
			number, ok := event.Data[key].(json.Number)
			value, err := number.Int64()
			if !ok || err != nil || value < 0 || value > 2147483647 {
				return nil, fmt.Errorf("event %d needs a Word integer %s", i+1, key)
			}
			coordinates[key] = value
		}
		if coordinates["end"] < coordinates["start"] || coordinates["story_type"] == 0 {
			return nil, fmt.Errorf("event %d has invalid range", i+1)
		}
		if _, ok := event.Data["text"].(string); !ok {
			return nil, fmt.Errorf("event %d lacks exact range text", i+1)
		}
		switch event.Kind {
		case "read", "plan", "skip":
		case "write":
			if property, ok := event.Data["property"].(string); !ok || property == "" {
				return nil, fmt.Errorf("write %d lacks property", i+1)
			}
			if _, ok := event.Data["requested"]; !ok {
				return nil, fmt.Errorf("write %d lacks requested value", i+1)
			}
			if event.Operation == "" || used[event.Operation] {
				return nil, fmt.Errorf("invalid/reused write operation at event %d", i+1)
			}
			pending[event.Operation], used[event.Operation] = event, true
		case "result":
			before, ok := pending[event.Operation]
			if !ok || before.Document != event.Document || before.Story != event.Story || before.Stage != event.Stage {
				return nil, fmt.Errorf("unpaired write result at event %d", i+1)
			}
			errorNumber, ok := event.Data["error"].(json.Number)
			_, numberErr := errorNumber.Int64()
			if !ok || numberErr != nil {
				return nil, fmt.Errorf("write result %d lacks error outcome", i+1)
			}
			delete(pending, event.Operation)
		default:
			return nil, fmt.Errorf("unknown trace event kind %q", event.Kind)
		}
	}
	if len(pending) != 0 {
		return nil, fmt.Errorf("trace has %d unfinished writes", len(pending))
	}
	return &trace, nil
}

func captureTrace(directory, name string) (*TraceEvidence, error) {
	path, err := project.Under(directory, name)
	if err != nil {
		return nil, err
	}
	raw, err := project.Read(directory, name)
	if err != nil {
		return nil, err
	}
	evidence := &TraceEvidence{File: path, SHA256: office.Hash(raw)}
	trace, err := decodeTrace(raw)
	if err != nil {
		return evidence, err
	}
	evidence.Events, evidence.Coverage = len(trace.Events), trace.Coverage
	var readable strings.Builder
	fmt.Fprintf(&readable, "# Ordered native trace\n\nRaw SHA-256: `%s`\n\nCoverage is declared by instrumentation, not a guarantee of exhaustive macro coverage.\nLive-anchor touches are interaction candidates, not proof of causality.\nThis report is not a DOCX byte-parity claim.\n\n", evidence.SHA256)
	for _, event := range trace.Events {
		fmt.Fprintf(&readable, "## Event %d\n\n", event.Sequence)
		data, _ := json.MarshalIndent(event, "", "  ")
		for _, line := range strings.Split(string(data), "\n") {
			fmt.Fprintf(&readable, "    %s\n", line)
		}
		readable.WriteString("\n")
	}
	if err := project.Write(directory, strings.TrimSuffix(name, ".json")+".md", []byte(readable.String()), ""); err != nil {
		return evidence, err
	}
	return evidence, nil
}

func readTrace(evidence *TraceEvidence) (*SerialTrace, error) {
	if evidence == nil {
		return nil, fmt.Errorf("serial trace evidence missing")
	}
	raw, err := project.Read(filepath.Dir(evidence.File), filepath.Base(evidence.File))
	if err != nil {
		return nil, err
	}
	if office.Hash(raw) != evidence.SHA256 {
		return nil, fmt.Errorf("serial trace SHA256 mismatch")
	}
	trace, err := decodeTrace(raw)
	if err != nil {
		return nil, err
	}
	if evidence.Events != len(trace.Events) || !bytes.Equal(project.JSON(evidence.Coverage), project.JSON(trace.Coverage)) {
		return nil, fmt.Errorf("serial trace summary mismatch")
	}
	return trace, nil
}

// Exact decoded event values are compared; no coordinates, text, formatting,
// errors or project-supplied fields are normalized or silently excluded.
func compareTraces(a, b *TraceEvidence) (*Difference, error) {
	x, err := readTrace(a)
	if err != nil {
		return nil, err
	}
	y, err := readTrace(b)
	if err != nil {
		return nil, err
	}
	if !bytes.Equal(project.JSON(x.Coverage), project.JSON(y.Coverage)) {
		return &Difference{Step: "serial trace", Path: "/coverage", Error: "trace coverage differs"}, nil
	}
	for i := 0; i < len(x.Events) || i < len(y.Events); i++ {
		var before, after any
		if i < len(x.Events) {
			before = x.Events[i]
		}
		if i < len(y.Events) {
			after = y.Events[i]
		}
		ab, _ := json.Marshal(before)
		bb, _ := json.Marshal(after)
		if !bytes.Equal(ab, bb) {
			return &Difference{Step: "serial trace", Path: fmt.Sprintf("/events/%d", i), Baseline: before, Candidate: after, Error: "first ordered event difference"}, nil
		}
	}
	return nil, nil
}
