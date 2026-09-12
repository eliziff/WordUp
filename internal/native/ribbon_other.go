//go:build !windows || (!amd64 && !arm64)

package native

func ValidateRibbon(data []byte) (map[string]any, error) {
	return map[string]any{"valid": false, "available": false, "reason": "Full RibbonX schema validation currently requires Windows MSXML 6.0"}, nil
}
