package agent

import (
	"encoding/json"
	"reflect"
	"strings"
	"testing"
)

func TestToolSchemasStayScopedToParameters(t *testing.T) {
	fields := map[string]bool{}
	typ := reflect.TypeOf(Parameters{})
	for i := 0; i < typ.NumField(); i++ {
		fields[strings.Split(typ.Field(i).Tag.Get("json"), ",")[0]] = true
	}
	seen := map[string]bool{}
	for _, tool := range Tools() {
		if seen[tool.Name] {
			t.Fatalf("duplicate tool %s", tool.Name)
		}
		seen[tool.Name] = true
		props := tool.InputSchema["properties"].(map[string]any)
		for name, schema := range props {
			if !fields[name] || schema == nil {
				t.Fatalf("%s has invalid property %s", tool.Name, name)
			}
		}
		if required, ok := tool.InputSchema["required"].([]string); ok {
			for _, name := range required {
				if props[name] == nil {
					t.Fatalf("missing required schema %s.%s", tool.Name, name)
				}
			}
		}
		if tool.Name != "native.call" && props["operation"] != nil {
			t.Fatalf("unrelated native schema in %s", tool.Name)
		}
		if tool.Name == "xml.patch" {
			if len(props) != 3 || len(tool.InputSchema["required"].([]string)) != 3 {
				t.Fatal("unguarded patch schema")
			}
		}
	}
	raw, err := json.Marshal(Tools())
	if err != nil {
		t.Fatal(err)
	}
	if len(raw) > 100000 {
		t.Fatalf("tool manifest unexpectedly bloated: %d bytes", len(raw))
	}
}

func BenchmarkToolManifest(b *testing.B) {
	raw, err := json.Marshal(Tools())
	if err != nil {
		b.Fatal(err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := json.Marshal(Tools()); err != nil {
			b.Fatal(err)
		}
	}
	b.ReportMetric(float64(len(raw)), "manifest-bytes")
}
