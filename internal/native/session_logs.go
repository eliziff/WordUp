package native

import (
	"encoding/base64"
	"encoding/binary"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf16"
)

// Only inspect the owned staging directory, never the user's global temp files.
func sessionLogs(directory string) map[string]any {
	logs := []any{}
	result := map[string]any{"files": logs, "directory": directory, "truncated": false}
	entries, err := os.ReadDir(directory)
	if err != nil {
		result["error"] = err.Error()
		return result
	}
	for _, entry := range entries {
		if entry.IsDir() || entry.Type()&os.ModeSymlink != 0 || !strings.EqualFold(filepath.Ext(entry.Name()), ".log") {
			continue
		}
		if len(logs) == 8 {
			result["truncated"] = true
			break
		}
		item := map[string]any{"name": entry.Name()}
		logs = append(logs, item)
		f, err := os.Open(filepath.Join(directory, entry.Name()))
		if err != nil {
			item["error"] = err.Error()
			continue
		}
		b, err := io.ReadAll(io.LimitReader(f, 262145))
		f.Close()
		if err != nil {
			item["error"] = err.Error()
			continue
		}
		item["truncated"] = len(b) > 262144
		if len(b) > 262144 {
			b = b[:262144]
		}
		item["base64"] = base64.StdEncoding.EncodeToString(b)
		text := string(b)
		if len(b) >= 2 && b[0] == 255 && b[1] == 254 {
			units := make([]uint16, 0, (len(b)-2)/2)
			for i := 2; i+1 < len(b); i += 2 {
				units = append(units, binary.LittleEndian.Uint16(b[i:]))
			}
			text = string(utf16.Decode(units))
		}
		item["text"] = text
	}
	result["files"] = logs
	return result
}
