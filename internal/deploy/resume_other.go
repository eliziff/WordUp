//go:build !windows

package deploy

func lockActivation(file string) (func(), error) { return func() {}, nil }
func registerResume(file, exe string) error      { return nil }
func clearResume(file string) error              { return nil }
