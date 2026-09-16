package kernels

import (
	"encoding/csv"
	"errors"
	"io"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

var (
	// Python's \s in [^\s]+ is Unicode whitespace; spelled out for RE2.
	originalURLRe = regexp.MustCompile(`^https?://[^\t\n\x0B\x0C\r \x1C-\x1F\x{85}\p{Z}]+$`)
	permaTargetRe = regexp.MustCompile(`^https://perma\.cc/[A-Z0-9]{4}-[A-Z0-9]{4}$`)
)

// PermaMapping parses the illustrative explicit CSV dialect
// submitted_url,status,perma_url. Column names are this kernel's import
// contract, not a claim about every Perma export version. Full RFC quoted
// records, duplicate conflicts and formula-like payloads are parsed as data;
// there is no spreadsheet execution or network activity.
func PermaMapping(csvText string) (map[string]string, error) {
	reader := csv.NewReader(strings.NewReader(strings.TrimPrefix(csvText, string(rune(0xFEFF)))))
	reader.FieldsPerRecord = -1
	reader.LazyQuotes = true
	header, err := reader.Read()
	if err != nil {
		return nil, errors.New("CSV schema mismatch")
	}
	required := []string{"submitted_url", "status", "perma_url"}
	index := map[string]int{}
	for i, name := range header {
		if _, dup := index[name]; dup {
			return nil, errors.New("CSV schema mismatch")
		}
		index[name] = i
	}
	for _, name := range required {
		if _, ok := index[name]; !ok {
			return nil, errors.New("CSV schema mismatch")
		}
	}
	result := map[string]string{}
	for {
		row, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, errors.New("malformed CSV record")
		}
		if len(row) != len(header) {
			return nil, errors.New("malformed CSV record")
		}
		if row[index["status"]] != "success" {
			continue
		}
		src, dst := row[index["submitted_url"]], row[index["perma_url"]]
		if !originalURLRe.MatchString(src) {
			return nil, errors.New("invalid original URL")
		}
		if !permaTargetRe.MatchString(dst) {
			return nil, errors.New("invalid Perma target")
		}
		if existing, ok := result[src]; ok && existing != dst {
			return nil, errors.New("conflicting exact URL mapping")
		}
		result[src] = dst
	}
	return result, nil
}

// PermaTokenChange replaces one caller-isolated full URL token, not a
// substring search. ok is false when the token has no mapping.
func PermaTokenChange(token string, mappings map[string]string) (string, bool, error) {
	target, ok := mappings[token]
	if !ok {
		return "", false, nil
	}
	if !permaTargetRe.MatchString(target) {
		return "", false, errors.New("mapping target is not a supported Perma URL")
	}
	return "[" + strings.TrimPrefix(target, "https://") + "]", true, nil
}

// Range is an inclusive numeric locator range; a single pinpoint has Lo == Hi.
type Range struct{ Lo, Hi int }

// ResolvedCitation is a fully supplied citation identity for the Ibid
// serializer. A nil Locator means source-wide, NOT inherited.
type ResolvedCitation struct {
	SourceID       string
	EditionID      string
	LocatorBasis   string
	Locator        []Range
	Qualifications []string
}

func rangesEqual(a, b []Range) bool {
	if (a == nil) != (b == nil) || len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

// IbidPayload outputs tokens only after identity/version and locator
// semantics are supplied. ok=false means "cannot prove equivalent under this
// deliberately limited serializer". There is no string-similarity fallback
// and no guessing of omitted pinpoints.
func IbidPayload(previous, current ResolvedCitation, previousNoteSingleSource, adjacent bool) (string, bool) {
	if !adjacent || !previousNoteSingleSource {
		return "", false
	}
	// Empty identities are UNKNOWN, not a shared source. Validate before
	// equality: equality of two malformed locators must not license omission.
	for _, item := range []ResolvedCitation{previous, current} {
		if strings.TrimSpace(item.SourceID) == "" || strings.TrimSpace(item.EditionID) == "" {
			return "", false
		}
		if item.LocatorBasis != "page" && item.LocatorBasis != "paragraph" {
			return "", false
		}
		if item.Locator != nil {
			if len(item.Locator) == 0 {
				return "", false
			}
			for _, r := range item.Locator {
				if r.Lo < 1 || r.Hi < r.Lo {
					return "", false
				}
			}
		}
	}
	if previous.SourceID != current.SourceID || previous.EditionID != current.EditionID || previous.LocatorBasis != current.LocatorBasis {
		return "", false
	}
	if len(current.Qualifications) > 0 || len(previous.Qualifications) > 0 {
		return "", false
	}
	if current.Locator == nil && previous.Locator != nil {
		return "", false
	}
	if rangesEqual(current.Locator, previous.Locator) {
		return "Ibid", true
	}
	if current.Locator == nil {
		return "Ibid", true
	}
	var out []string
	multiple := len(current.Locator) > 1
	for _, r := range current.Locator {
		if r.Lo < 1 || r.Hi < r.Lo {
			return "", false
		}
		if r.Lo == r.Hi {
			out = append(out, strconv.Itoa(r.Lo))
			continue
		}
		multiple = true
		end := strconv.Itoa(r.Hi)
		if current.LocatorBasis == "page" {
			shortened, err := ShortenPageRange(strconv.Itoa(r.Lo), end)
			if err != nil {
				return "", false
			}
			end = shortened
		}
		out = append(out, strconv.Itoa(r.Lo)+"–"+end)
	}
	label := ""
	if current.LocatorBasis != "page" {
		label = "para "
		if multiple {
			label = "paras "
		}
	}
	return "Ibid at " + label + strings.Join(out, ", "), true
}

// Italic is a phrase-italics intention over a code-point range.
type Italic struct {
	Start, End int
	Italic     bool
}

// PhraseItalics returns intentions for the longest complete matching phrase
// at each position; results are intentions, not automatic deitalicization.
// A false rule suppresses a shorter positive match; it does NOT authorize
// removing independently meaningful author/source italics from a Word run.
func PhraseItalics(text string, rules map[string]bool, extra []Span) ([]Italic, error) {
	runes := []rune(text)
	guards, err := protection(runes, text, extra)
	if err != nil {
		return nil, err
	}
	phrases := make([]string, 0, len(rules))
	for p := range rules {
		phrases = append(phrases, p)
	}
	sort.Slice(phrases, func(i, j int) bool {
		if len([]rune(phrases[i])) != len([]rune(phrases[j])) {
			return len([]rune(phrases[i])) > len([]rune(phrases[j]))
		}
		return phrases[i] < phrases[j]
	})
	var out []Italic
	for i := 0; i < len(runes); {
		found := false
		for _, phrase := range phrases {
			j := i + len([]rune(phrase))
			if j > len(runes) || pytext.Casefold(string(runes[i:j])) != pytext.Casefold(phrase) {
				continue
			}
			if (i > 0 && pytext.IsWord(runes[i-1])) || (j < len(runes) && pytext.IsWord(runes[j])) {
				continue
			}
			blocked := false
			for _, g := range guards {
				if g.Overlaps(i, j) {
					blocked = true
					break
				}
			}
			if blocked {
				continue
			}
			out = append(out, Italic{i, j, rules[phrase]})
			i = j
			found = true
			break
		}
		if !found {
			i++
		}
	}
	return out, nil
}
