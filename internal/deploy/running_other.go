//go:build !windows && !darwin

package deploy

func WordRunning() (bool, error) { return false, nil }
