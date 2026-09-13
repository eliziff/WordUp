//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"errors"
	"testing"
	"time"
)

func TestOfficeWorkerQueuedCallHonorsCancellation(t *testing.T) {
	if err := lockOfficeWorkers(context.Background()); err != nil {
		t.Fatal(err)
	}
	defer unlockOfficeWorkers()
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan error, 1)
	go func() {
		_, err := callOfficeWorker(ctx, "must-not-launch.exe", nil)
		done <- err
	}()
	cancel()
	select {
	case err := <-done:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("queued cancellation: %v", err)
		}
	case <-time.After(time.Second):
		t.Fatal("queued call ignored cancellation while another worker held the gate")
	}
}

func TestOfficeWorkerCancelledCallDoesNotLaunch(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err := callOfficeWorker(ctx, "must-not-launch.exe", nil)
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("cancelled call attempted launch: %v", err)
	}
}

func TestOfficeWorkerRejectsInvalidResponseObjects(t *testing.T) {
	for _, input := range []string{"null", "[]", "true", "42", `"text"`, "{"} {
		if _, err := decodeOfficeWorkerResponse([]byte(input)); err == nil {
			t.Errorf("accepted invalid helper response %q", input)
		}
	}
	for _, input := range []string{`{}`, `{"valid":false}`, `{"error":"failed"}`} {
		value, err := decodeOfficeWorkerResponse([]byte(input))
		if err != nil || value == nil {
			t.Errorf("rejected object %q: %v", input, err)
		}
	}
}
