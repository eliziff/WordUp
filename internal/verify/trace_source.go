package verify

import _ "embed"

// TraceRecorderSource is a standalone opt-in Word VBA module. Installing it
// does not instrument a project: callers explicitly observe their own paths.
//
//go:embed vba/WordUp_Trace.bas
var TraceRecorderSource string
