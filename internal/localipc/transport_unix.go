//go:build !windows

package localipc

import (
	"context"
	"io"
	"net"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"syscall"
)

type listener struct {
	net.Listener
	dir string
}

func (l *listener) Address() string                     { return l.Addr().String() }
func (l *listener) Close() error                        { e := l.Listener.Close(); os.RemoveAll(l.dir); return e }
func (l *listener) Accept() (io.ReadWriteCloser, error) { return l.Listener.Accept() }
func listen() (*listener, error) {
	d, e := os.MkdirTemp("", "ww-ipc-")
	if e != nil {
		return nil, e
	}
	p := filepath.Join(d, "socket")
	l, e := net.Listen("unix", p)
	if e != nil {
		os.RemoveAll(d)
		return nil, e
	}
	if e = os.Chmod(p, 0600); e != nil {
		l.Close()
		os.RemoveAll(d)
		return nil, e
	}
	return &listener{l, d}, nil
}
func validAddress(s string) bool {
	return filepath.IsAbs(s) && filepath.Base(s) == "socket" && strings.HasPrefix(filepath.Base(filepath.Dir(s)), "ww-ipc-")
}
func dial(ctx context.Context, address string) (io.ReadWriteCloser, error) {
	return (&net.Dialer{}).DialContext(ctx, "unix", address)
}
func detach(cmd *exec.Cmd) { cmd.SysProcAttr = &syscall.SysProcAttr{Setsid: true} }
