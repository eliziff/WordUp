package deploy

import (
	"errors"
	"fmt"
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"os/exec"
	"path/filepath"
)

func Spawn(plan string) error {
	plan, e := filepath.Abs(plan)
	if e != nil {
		return e
	}
	exe, e := os.Executable()
	if e != nil {
		return e
	}
	if e = registerResume(plan, exe); e != nil {
		return e
	}
	log, e := os.OpenFile(filepath.Join(filepath.Dir(plan), "activation.log"), os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0600)
	if e != nil {
		return e
	}
	defer log.Close()
	cmd := exec.Command(exe, "__activate", plan)
	detach(cmd)
	cmd.Stdout, cmd.Stderr = log, log
	if e = cmd.Start(); e != nil {
		return e
	}
	return cmd.Process.Release()
}
func RecordFailure(file string, err error) {
	if errors.Is(err, ErrActivationBusy) || errors.Is(err, ErrWordRunning) {
		return
	}
	_ = clearResume(file)
	b, e := project.Read(filepath.Dir(file), filepath.Base(file))
	if e != nil {
		return
	}
	p := &Plan{}
	if project.ReadJSON(b, p) != nil {
		return
	}
	if p.State == "pending" {
		p.State = "blocked"
		p.Error = fmt.Sprint(err)
		_ = project.AtomicWrite(file, project.JSON(p))
	}
}
