package office

import (
	"strings"
	"testing"
)

func TestProjectComponentCollisionFailsBeforeWord(t *testing.T) {
	v := NewVBA("Atelier")
	for _, name := range []string{"Atelier", "ATELIER"} {
		_, err := v.Rewrite([]Module{{Name: name, Kind: "standard", Source: "Option Explicit\n"}}, nil)
		if err == nil || !strings.Contains(err.Error(), "conflicts with VBA project") {
			t.Fatalf("missing actionable collision error: %v", err)
		}
	}
	if _, err := v.Rewrite([]Module{{Name: "AtelierActions", Kind: "standard", Source: "Option Explicit\n"}}, nil); err != nil {
		t.Fatal(err)
	}
}
