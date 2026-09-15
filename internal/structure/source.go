// Package structure supplies an editable, standalone Word VBA detection core.
package structure

import _ "embed"

//go:embed WordUpStructure.bas
var Source string

const ContractVersion = "1.2.9"
