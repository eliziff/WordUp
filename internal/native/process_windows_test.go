//go:build windows && (amd64 || arm64)

package native

import (
	"errors"
	"fmt"
	"syscall"
	"testing"
)

func TestTerminationCauseSurvivesClosedSessionAndPendingCalls(t *testing.T) {
	for _, code := range []string{"macro_deadline", "worker_exited", "word_exited"} {
		t.Run(code, func(t *testing.T) {
			cause := fault(Fail(code, "original cause", map[string]any{"task": "edit"}))
			response := make(chan Response, 1)
			h := &localHost{termination: cause, pending: map[uint64]chan Response{1: response}}
			if !errors.Is(h.closedFault(), cause) {
				t.Fatal("closed session lost its termination cause")
			}
			h.breakPending(Fail("session_closed", "cleanup", nil))
			if got := <-response; got.Error != cause {
				t.Fatalf("pending call lost cause: %#v", got.Error)
			}
		})
	}
	if got := fault((&localHost{}).closedFault()); got.Code != "session_closed" {
		t.Fatalf("ordinary close misclassified: %#v", got)
	}
}

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

func TestCallerLaunchTokenFailureExplainsRecovery(t *testing.T) {
	f := fault(callerLaunchTokenFailure(map[string]any{"create_error": "1312"}))
	if f.Code != "native_launch_token_restricted" {
		t.Fatalf("wrong classification: %#v", f)
	}
	details := f.Details.(map[string]any)
	if details["session_restart_helpful"] != false || details["recovery"] == "" {
		t.Fatalf("missing actionable recovery: %#v", details)
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

// The named constants must carry the documented Win32 values. An earlier
// revision swapped two of them, and a name-only assertion let Word launches
// fail for a whole day with ERROR_NO_SUCH_LOGON_SESSION.
func TestProcessCreationFlagValuesMatchWin32(t *testing.T) {
	for _, c := range []struct {
		name string
		got  uint32
		want uint32
	}{
		{"CREATE_SUSPENDED", createSuspended, 0x00000004},
		{"CREATE_NO_WINDOW", createNoWindow, 0x08000000},
		{"CREATE_BREAKAWAY_FROM_JOB", createBreakaway, 0x01000000},
		{"EXTENDED_STARTUPINFO_PRESENT", extendedStartupInfo, 0x00080000},
	} {
		if c.got != c.want {
			t.Fatalf("%s = %#x, want %#x", c.name, c.got, c.want)
		}
	}
}
