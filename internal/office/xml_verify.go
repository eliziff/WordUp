package office

import "fmt"

// VerifyXML compares the expected artifact directly, without a manifest or
// annotation roundtrip. Equality covers the whole input; no ignored content.
func VerifyXML(expected, actual []byte, mode string) (map[string]any, error) {
	if mode == "" {
		mode = "exact"
	}
	if mode != "exact" && mode != "semantic" {
		return nil, fmt.Errorf("comparison must be exact or semantic")
	}
	for _, data := range [][]byte{expected, actual} {
		if _, err := XMLSpans(data); err != nil {
			return nil, err
		}
	}
	report, err := CompareXML(expected, actual, XMLComparePolicy{})
	if err != nil {
		return nil, err
	}
	matched := report["byte_identical"] == true
	if mode == "semantic" {
		matched = report["equal"] == true
	}
	report["comparison"] = mode
	report["matches_expected"] = matched
	if !matched {
		// Include an exact offset even for serialization-only differences.
		i := 0
		for i < len(expected) && i < len(actual) && expected[i] == actual[i] {
			i++
		}
		report["first_different_byte"] = i
		return report, fmt.Errorf("actual XML differs from expected XML (%s comparison)", mode)
	}
	return report, nil
}
