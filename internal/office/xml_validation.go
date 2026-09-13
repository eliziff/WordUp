package office

import (
	"encoding/binary"
	"fmt"
	"regexp"
	"unicode/utf16"
)

var utf16Declaration = regexp.MustCompile(`(?i)(encoding\s*=\s*["'])utf-16(?:le|be)?(["'])`)

// Validation needs names and attributes, not edit offsets. Decode a temporary
// view of UTF-16 XML here; preserve original part bytes and UTF-8 edit spans.
func validationXMLSpans(b []byte) ([]XMLSpan, error) {
	if len(b) < 2 || !(b[0] == 255 && b[1] == 254 || b[0] == 254 && b[1] == 255) {
		return XMLSpans(b)
	}
	if len(b) > Limit {
		return nil, fmt.Errorf("XML budget exceeded")
	}
	if len(b)%2 != 0 {
		return nil, fmt.Errorf("truncated UTF-16 XML")
	}
	var order binary.ByteOrder = binary.LittleEndian
	if b[0] == 254 {
		order = binary.BigEndian
	}
	units := make([]uint16, 0, (len(b)-2)/2)
	for i := 2; i < len(b); i += 2 {
		units = append(units, order.Uint16(b[i:i+2]))
	}
	for i := 0; i < len(units); i++ {
		if units[i] >= 0xD800 && units[i] <= 0xDBFF {
			if i+1 >= len(units) || units[i+1] < 0xDC00 || units[i+1] > 0xDFFF {
				return nil, fmt.Errorf("unpaired UTF-16 surrogate")
			}
			i++
		} else if units[i] >= 0xDC00 && units[i] <= 0xDFFF {
			return nil, fmt.Errorf("unpaired UTF-16 surrogate")
		}
	}
	decoded := []byte(string(utf16.Decode(units)))
	if len(decoded) > Limit {
		return nil, fmt.Errorf("XML budget exceeded")
	}
	// Only the declaration may declare encoding; do not rewrite document content.
	for i := 0; i+1 < len(decoded) && i < 512; i++ {
		if decoded[i] == '?' && decoded[i+1] == '>' {
			head := utf16Declaration.ReplaceAll(decoded[:i+2], []byte("${1}utf-8${2}"))
			decoded = append(head, decoded[i+2:]...)
			break
		}
	}
	return XMLSpans(decoded)
}
