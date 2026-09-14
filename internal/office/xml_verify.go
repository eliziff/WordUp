package office

import "fmt"

// XMLInput selects an XML part from an OPC package when part is non-empty.
// Raw XML remains valid input, which lets an expected XML file be compared
// directly with a package part without an extraction fixture.
func XMLInput(data []byte, part string) ([]byte, error) {
	if part == "" {
		return data, nil
	}
	if !SafePart(part) {
		return nil, fmt.Errorf("unsafe XML package part %q", part)
	}
	pkg, packageErr := ReadPackage(data)
	if packageErr == nil {
		selected, ok := pkg.Files[part]
		if !ok {
			return nil, fmt.Errorf("XML package part %q not found", part)
		}
		return selected, nil
	}
	// A raw expected XML file has no package part to select; use it as-is.
	// Validate it here so a malformed raw file does not produce a misleading
	// comparison error later.
	if _, xmlErr := XMLSpans(data); xmlErr == nil {
		return data, nil
	}
	return nil, fmt.Errorf("select XML package part %q: %w", part, packageErr)
}

// VerifyXML compares the expected artifact directly, without a manifest or
// secondary representation. Equality covers the whole input; no ignored
// content.
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
