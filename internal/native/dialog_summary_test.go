package native

import (
	"reflect"
	"testing"
)

func TestDialogSummaryKeepsMessagesAndButtons(t *testing.T) {
	tree := map[string]any{"name": "Visual Basic", "role": 9, "children": []any{
		map[string]any{"role": int32(41), "name": "Run-time error 515: budget exceeded"},
		map[string]any{"role": float64(41), "name": "Run-time error 515: budget exceeded"},
		map[string]any{"role": 43, "name": "End"},
		map[string]any{"role": 9, "children": []any{map[string]any{"role": 43, "name": "Debug"}}},
	}}
	got := dialogSummary(tree)
	if !reflect.DeepEqual(got["messages"], []string{"Run-time error 515: budget exceeded"}) || !reflect.DeepEqual(got["buttons"], []string{"End", "Debug"}) {
		t.Fatal(got)
	}
}
