package office

import (
	"bytes"
	"fmt"
)

const flatOPCNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage"

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
	spans, xmlErr := XMLSpans(data)
	if xmlErr == nil {
		if len(spans) > 0 && spans[0].Name.Space == flatOPCNamespace && spans[0].Name.Local == "package" {
			want := "/" + part
			for i, span := range spans {
				if span.Name.Space != flatOPCNamespace || span.Name.Local != "part" || span.Attribute(flatOPCNamespace, "name") != want {
					continue
				}
				for _, dataSpan := range spans[i+1:] {
					if dataSpan.Start >= span.End {
						break
					}
					if dataSpan.Name.Space != flatOPCNamespace || dataSpan.Name.Local != "xmlData" || dataSpan.Depth != span.Depth+1 {
						continue
					}
					for _, xmlSpan := range spans[i+1:] {
						if xmlSpan.Start >= dataSpan.End {
							break
						}
						if xmlSpan.Depth == dataSpan.Depth+1 {
							return data[xmlSpan.Start:xmlSpan.End], nil
						}
					}
					return nil, fmt.Errorf("Flat OPC part %q has no XML data", part)
				}
				return nil, fmt.Errorf("Flat OPC part %q not found", part)
			}
			return nil, fmt.Errorf("Flat OPC part %q not found", part)
		}
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
	// Exact bytes are the common parity-fixture success path. Validate once,
	// then return the same evidence shape without tokenizing both copies.
	if bytes.Equal(expected, actual) {
		if _, err := XMLSpans(expected); err != nil {
			return nil, err
		}
		hash := Hash(expected)
		return map[string]any{
			"equal": true, "reference_sha256": hash, "candidate_sha256": hash,
			"byte_identical": true, "comparison": mode, "matches_expected": true,
			"policy": XMLComparePolicy{},
		}, nil
	}
	// CompareXML performs the same strict XML validation while tokenizing the
	// two inputs. Do not walk both documents once here and then immediately
	// walk them again in the comparator: mismatch diagnostics are the common
	// edit-loop failure path, and the duplicate parse was measurable there.
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
