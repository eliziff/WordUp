//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"syscall"
	"testing"
)

func TestNoLogonSessionErrorClassification(t *testing.T) {
	if !noLogonSessionError(fmt.Errorf("CreateProcessW: %w", syscall.Errno(1312))) {
		t.Fatal("1312 must trigger the restricted-session fallback")
	}
	if !noLogonSessionError(fmt.Errorf("The specified logon session does not exist")) {
		t.Fatal("the Windows error text must trigger the restricted-session fallback")
	}
	if noLogonSessionError(syscall.Errno(5)) {
		t.Fatal("unrelated Windows errors must not trigger the fallback")
	}
}

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
		if flags&(createSuspended|createNoWindow) != createSuspended|createNoWindow {
			t.Fatalf("required creation flags were lost: %#x", flags)
		}
		if flags&extendedStartupInfo == 0 {
			t.Fatalf("extended startup info flag was lost: %#x", flags)
		}
	}
}
