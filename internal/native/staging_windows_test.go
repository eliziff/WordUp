//go:build windows && (amd64 || arm64)

package native

import (
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"testing"
)

func TestStagingStillRejectsChangedBytes(t *testing.T) {
	source := filepath.Join(t.TempDir(), "input.docx")
	pkg := office.BlankPackage()
	data, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(source, data, 0600); err != nil {
		t.Fatal(err)
	}
	h := &wordHost{cfg: hostConfig{Options: Options{Directory: t.TempDir()}}, staged: map[string]string{}}
	staged, err := h.stage(source)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := h.stage(source); err != nil {
		t.Fatal("unchanged staged bytes not reusable", err)
	}
	// Word may retain a read lock after the document handle is unloaded.
	path, err := syscall.UTF16PtrFromString(staged)
	if err != nil {
		t.Fatal(err)
	}
	lock, err := syscall.CreateFile(path, syscall.GENERIC_READ, syscall.FILE_SHARE_READ, nil, syscall.OPEN_EXISTING, 0, 0)
	if err != nil {
		t.Fatal(err)
	}
	delete(h.staged, strings.ToLower(staged))
	_, err = h.stage(source)
	syscall.CloseHandle(lock)
	if err != nil {
		t.Fatal("unchanged unloaded template with retained read lock not reusable", err)
	}
	if err := os.WriteFile(staged, []byte("modified"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := h.stage(source); err == nil {
		t.Fatal("changed staged bytes accepted")
	}
}
