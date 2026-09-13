package office

import (
	"bytes"
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

func TestRewritePreservesUnchangedModuleStream(t *testing.T) {
	modules := []Module{
		{Name: "First", Kind: "standard", Source: "Option Explicit\nPublic Sub One(): End Sub\n"},
		{Name: "Second", Kind: "standard", Source: "Option Explicit\nPublic Sub Two(): End Sub\n"},
	}
	raw, err := NewVBA("Proof").Rewrite(modules, nil)
	if err != nil {
		t.Fatal(err)
	}
	before, err := ReadVBA(raw)
	if err != nil {
		t.Fatal(err)
	}
	preserved := map[string][]byte{}
	for _, name := range []string{"VBA/First", "VBA/dir", "PROJECT", "PROJECTwm"} {
		preserved[name], err = before.CFB.Stream(name)
		if err != nil {
			t.Fatal(err)
		}
	}
	modules[1].Source += "' changed\n"
	raw, err = before.Rewrite(modules, nil)
	if err != nil {
		t.Fatal(err)
	}
	after, err := ReadVBA(raw)
	if err != nil {
		t.Fatal(err)
	}
	for name, want := range preserved {
		got, streamErr := after.CFB.Stream(name)
		if streamErr != nil {
			t.Fatal(streamErr)
		}
		if !bytes.Equal(want, got) {
			t.Fatalf("editing one module rewrote unchanged stream %s", name)
		}
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
