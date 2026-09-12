//go:build windows && (amd64 || arm64)

package deploy

import (
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"syscall"
	"testing"
	"time"
	"unsafe"

	"github.com/eliziff/WordUp/internal/example"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestMain(m *testing.M) {
	if len(os.Args) > 1 && os.Args[1] == "__host" {
		if native.HostMain(os.Args[2:]) != nil {
			os.Exit(1)
		}
		return
	}
	if len(os.Args) == 3 && os.Args[1] == "__activate" {
		if Watch(context.Background(), os.Args[2]) != nil {
			os.Exit(1)
		}
		return
	}
	os.Exit(m.Run())
}

// Executes the registered login script, not an actual OS reboot.
func TestNativeLoginRecovery(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := t.TempDir()
	build, err := example.Studio(filepath.Join(root, "studio"))
	if err != nil {
		t.Fatal(err)
	}
	r, err := verify.Run(context.Background(), build.Artifact, example.WindowsSuite(), nil, true)
	if err != nil {
		t.Fatal(err)
	}
	proof := filepath.Join(root, "proof.json")
	if err = os.WriteFile(proof, project.JSON(r), 0600); err != nil {
		t.Fatal(err)
	}
	target := filepath.Join(root, "installed.dotm")
	if err = os.WriteFile(target, []byte("previous template"), 0600); err != nil {
		t.Fatal(err)
	}
	file, p, err := Prepare(root, build.Artifact, proof, target)
	if err != nil {
		t.Fatal(err)
	}
	exe, err := os.Executable()
	if err != nil {
		t.Fatal(err)
	}
	if err = registerResume(file, exe); err != nil {
		t.Fatal(err)
	}
	defer clearResume(file)
	var key syscall.Handle
	if err = syscall.RegOpenKeyEx(syscall.HKEY_CURRENT_USER, syscall.StringToUTF16Ptr(`Software\Microsoft\Windows\CurrentVersion\Run`), 0, syscall.KEY_READ, &key); err != nil {
		t.Fatal(err)
	}
	defer syscall.RegCloseKey(key)
	readRun := func() (string, error) {
		var value [4096]uint16
		size := uint32(len(value) * 2)
		err := syscall.RegQueryValueEx(key, syscall.StringToUTF16Ptr(activationName(file)), nil, nil, (*byte)(unsafe.Pointer(&value[0])), &size)
		return syscall.UTF16ToString(value[:]), err
	}
	wscript := filepath.Join(os.Getenv("SystemRoot"), "System32", "wscript.exe")
	script := filepath.Join(filepath.Dir(file), "resume.vbs")
	command, err := readRun()
	if err != nil || command != syscall.EscapeArg(wscript)+" //B //Nologo "+syscall.EscapeArg(script) {
		t.Fatalf("login command: %q %v", command, err)
	}
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	cmd := exec.CommandContext(ctx, wscript, "//B", "//Nologo", script)
	detach(cmd)
	if err = cmd.Run(); err != nil {
		t.Fatal(err)
	}
	for {
		data, _ := os.ReadFile(file)
		_ = project.ReadJSON(data, p)
		if p.State == "installed" {
			break
		}
		select {
		case <-ctx.Done():
			t.Fatal("login activation did not finish")
		case <-time.After(25 * time.Millisecond):
		}
	}
	installed, _ := os.ReadFile(target)
	if office.Hash(installed) != r.SHA256 {
		t.Fatal("installed bytes differ from native-tested artifact")
	}
	if _, err = readRun(); err != syscall.ERROR_FILE_NOT_FOUND {
		t.Fatalf("completed activation left login entry: %v", err)
	}
	if _, err = Restore(file); err != nil {
		t.Fatal(err)
	}
	previous, _ := os.ReadFile(target)
	if string(previous) != "previous template" {
		t.Fatal("restore differs")
	}
	t.Logf("fresh native assertions=%d; login script installed exact artifact; restore passed", r.Assertions)
}
