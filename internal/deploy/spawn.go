package deploy

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"wordwright.local/internal/project"
)

func Spawn(plan string) error {
	exe, e := os.Executable()
	if e != nil {
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
