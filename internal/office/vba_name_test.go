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

func TestNewClassGetsNativeClassIdentity(t *testing.T) {
	source := NewModuleSource(Module{Name: "ProofClass", Kind: "class", Source: "Option Explicit\n"}, "Proof")
	for _, attribute := range []string{
		`Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"`,
		"Attribute VB_TemplateDerived = False",
		"Attribute VB_Customizable = False",
	} {
		if !strings.Contains(source, attribute) {
			t.Fatalf("new class lacks %s", attribute)
		}
	}
}
