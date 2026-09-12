//go:build !windows

package office

import "fmt"

func platformDecode(b []byte, cp int) (string, error) {
	return "", fmt.Errorf("codepage %d requires the Windows codec backend; original bytes were not changed", cp)
}
func platformEncode(s string, cp int) ([]byte, error) {
	return nil, fmt.Errorf("codepage %d requires the Windows codec backend", cp)
}
