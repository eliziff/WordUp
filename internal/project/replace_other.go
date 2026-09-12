//go:build !windows

package project

import "os"

func replaceFile(from, to string) error { return os.Rename(from, to) }
