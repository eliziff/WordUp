package office

import (
	"fmt"
	"unicode/utf8"
)

var cp1252 = []rune("€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ")
var macRoman = []rune("ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ")

func Decode(b []byte, cp int) (string, error) {
	if cp == 65001 {
		if !utf8.Valid(b) {
			return "", fmt.Errorf("invalid UTF-8")
		}
		return string(b), nil
	}
	if cp == 1200 {
		if len(b)%2 != 0 {
			return "", fmt.Errorf("invalid UTF-16")
		}
		return utf16str(b), nil
	}
	var table []rune
	switch cp {
	case 1252:
		table = cp1252
	case 10000:
		table = macRoman
	default:
		return platformDecode(b, cp)
	}
	r := make([]rune, len(b))
	for i, v := range b {
		if v < 128 {
			r[i] = rune(v)
		} else {
			r[i] = table[int(v)-128]
		}
	}
	return string(r), nil
}
func Encode(s string, cp int) ([]byte, error) {
	if cp == 65001 {
		return []byte(s), nil
	}
	if cp == 1200 {
		return utf16bytes(s), nil
	}
	var table []rune
	switch cp {
	case 1252:
		table = cp1252
	case 10000:
		table = macRoman
	default:
		return platformEncode(s, cp)
	}
	out := []byte{}
	for _, r := range s {
		if r < 128 {
			out = append(out, byte(r))
			continue
		}
		found := false
		for i, v := range table {
			if v == r {
				out = append(out, byte(i+128))
				found = true
				break
			}
		}
		if !found {
			return nil, fmt.Errorf("U+%04X is not representable in VBA codepage %d; use ChrW or change project encoding deliberately", r, cp)
		}
	}
	return out, nil
}
