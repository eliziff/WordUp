//go:build !windows

package project

func fileChangeStamp(string) (int64, bool) { return 0, false }
