package agent

import (
	"context"
	"testing"
)

func TestImmediateRejectsInvalidInputWithoutStartingWord(t *testing.T) {
	for _, tc := range []struct {
		execute bool
		text    string
		timeout int
	}{
		{false, "? 1", 0}, {true, "Debug.Print 1", 0}, {true, "? ", 0},
		{true, "? 1\nEnd Function", 0}, {true, "? 1", 1800001},
	} {
		e := &Engine{Root: t.TempDir(), Execute: tc.execute}
		if _, err := e.Call(context.Background(), "vba.immediate", Parameters{Text: tc.text, TimeoutMS: tc.timeout}); err == nil {
			t.Fatalf("accepted invalid request: %+v", tc)
		}
		if e.host != nil {
			e.Close()
			t.Fatal("invalid request started Word")
		}
	}
}
