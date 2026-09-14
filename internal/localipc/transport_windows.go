//go:build windows

package localipc

import (
	"context"
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"io"
	"os"
	"os/exec"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"
)

var k = syscall.NewLazyDLL("kernel32.dll")
var createPipe = k.NewProc("CreateNamedPipeW")
var connectPipe = k.NewProc("ConnectNamedPipe")
var closeH = k.NewProc("CloseHandle")
var localFree = k.NewProc("LocalFree")
var a = syscall.NewLazyDLL("advapi32.dll")
var sddl = a.NewProc("ConvertStringSecurityDescriptorToSecurityDescriptorW")

type securityAttributes struct {
	Length     uint32
	Descriptor uintptr
	Inherit    int32
}
type listener struct {
	address string
	mu      sync.Mutex
	pending syscall.Handle
	closed  bool
}

func (l *listener) Address() string { return l.address }
func (l *listener) Close() error {
	l.mu.Lock()
	l.closed = true
	pending := l.pending
	l.mu.Unlock()
	if pending != 0 {
		name, _ := syscall.UTF16PtrFromString(l.address)
		h, err := syscall.CreateFile(name, syscall.GENERIC_READ|syscall.GENERIC_WRITE, 0, nil, syscall.OPEN_EXISTING, 0, 0)
		if err == nil {
			syscall.CloseHandle(h)
		}
	}
	return nil
}
func listen() (*listener, error) {
	b := make([]byte, 16)
	if _, e := rand.Read(b); e != nil {
		return nil, e
	}
	return &listener{address: `\\.\pipe\WordUp-` + hex.EncodeToString(b)}, nil
}
func validAddress(s string) bool {
	return strings.HasPrefix(s, `\\.\pipe\WordUp-`) && len(s) == len(`\\.\pipe\WordUp-`)+32
}
func (l *listener) Accept() (io.ReadWriteCloser, error) {
	token, e := syscall.OpenCurrentProcessToken()
	if e != nil {
		return nil, e
	}
	user, e := token.GetTokenUser()
	token.Close()
	if e != nil {
		return nil, e
	}
	sid, e := user.User.Sid.String()
	if e != nil {
		return nil, e
	}
	text, e := syscall.UTF16PtrFromString("D:P(A;;GA;;;" + sid + ")")
	if e != nil {
		return nil, e
	}
	var sd uintptr
	r, _, e := sddl.Call(uintptr(unsafe.Pointer(text)), 1, uintptr(unsafe.Pointer(&sd)), 0)
	if r == 0 {
		return nil, e
	}
	defer localFree.Call(sd)
	sa := securityAttributes{Length: uint32(unsafe.Sizeof(securityAttributes{})), Descriptor: sd}
	name, _ := syscall.UTF16PtrFromString(l.address)
	l.mu.Lock()
	if l.closed {
		l.mu.Unlock()
		return nil, os.ErrClosed
	}
	h, _, e := createPipe.Call(uintptr(unsafe.Pointer(name)), 3, 8, 32, 65536, 65536, 5000, uintptr(unsafe.Pointer(&sa)))
	if h == ^uintptr(0) {
		l.mu.Unlock()
		return nil, e
	}
	l.pending = syscall.Handle(h)
	l.mu.Unlock()
	ok, _, err := connectPipe.Call(h, 0)
	l.mu.Lock()
	if l.pending == syscall.Handle(h) {
		l.pending = 0
	}
	closed := l.closed
	l.mu.Unlock()
	if closed {
		closeH.Call(h)
		return nil, os.ErrClosed
	}
	if ok == 0 && err != syscall.Errno(535) {
		closeH.Call(h)
		return nil, err
	}
	return os.NewFile(h, l.address), nil
}
func dial(ctx context.Context, address string) (io.ReadWriteCloser, error) {
	name, e := syscall.UTF16PtrFromString(address)
	if e != nil {
		return nil, e
	}
	for {
		h, e := syscall.CreateFile(name, syscall.GENERIC_READ|syscall.GENERIC_WRITE, 0, nil, syscall.OPEN_EXISTING, 0, 0)
		if e == nil {
			return os.NewFile(uintptr(h), address), nil
		}
		if e != syscall.Errno(231) {
			return nil, fmt.Errorf("named pipe connection: %w", e)
		}
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-time.After(10 * time.Millisecond):
		}
	}
}
func detach(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{CreationFlags: 0x08000000 | 0x00004000} // CREATE_NO_WINDOW | BELOW_NORMAL_PRIORITY_CLASS
}
