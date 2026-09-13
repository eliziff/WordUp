package project

import (
	"crypto/sha256"
	"encoding/hex"
	"io"
	"os"
	"sync"
)

// Hash once per process; copied executables share identity, rebuilt ones do not.
var executableSHA256 = sync.OnceValue(func() string {
	path, err := os.Executable()
	if err != nil {
		return ""
	}
	file, err := os.Open(path)
	if err != nil {
		return ""
	}
	defer file.Close()
	hash := sha256.New()
	if _, err = io.Copy(hash, file); err != nil {
		return ""
	}
	return hex.EncodeToString(hash.Sum(nil))
})

func ExecutableSHA256() string { return executableSHA256() }
