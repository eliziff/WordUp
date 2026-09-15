//go:build windows && (amd64 || arm64)

package native

import "testing"

func TestProcessCreationFlagsKeepWordInOwnedJob(t *testing.T) {
	worker := processCreationFlags(1)
	if worker&createBreakaway == 0 {
		t.Fatal("owned worker must break away before job assignment")
	}
	word := processCreationFlags(0)
	if word&createBreakaway != 0 {
		t.Fatal("Word child must inherit the owned job")
	}
	for _, flags := range []uint32{worker, word} {
		if flags&(createNewProcessGroup|createSuspended|createNoWindow) != createNewProcessGroup|createSuspended|createNoWindow {
			t.Fatalf("required creation flags were lost: %#x", flags)
		}
	}
}
