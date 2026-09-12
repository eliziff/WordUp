package verify

import (
	"os"
	"path/filepath"
)

func createCandidate(source string, b []byte) (string, error) {
	d, e := os.MkdirTemp("", "wordup-candidate-")
	if e != nil {
		return "", e
	}
	p := filepath.Join(d, filepath.Base(source))
	if e = os.WriteFile(p, b, 0600); e != nil {
		os.RemoveAll(d)
		return "", e
	}
	return p, nil
}
func removeCandidate(file string) { _ = os.RemoveAll(filepath.Dir(file)) }
