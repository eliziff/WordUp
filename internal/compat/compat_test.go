package compat

import "testing"

func TestBranchReview(t *testing.T) {
	s := `#If Mac Then
Public Sub Test()
End Sub
#Else
Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal a As Long) As Long
#End If
`
	ds := Source("Module.bas", s, profiles[2])
	for _, d := range ds {
		if d.Rule == "windows-library-on-mac" {
			t.Fatal("inactive branch flagged")
		}
	}
	s = `#If Win64 Then
Private Declare Function Beep Lib "kernel32" () As Long
#End If`
	ds = Source("M.bas", s, profiles[2])
	found := false
	for _, d := range ds {
		found = found || d.Rule == "windows-library-on-mac"
	}
	if !found {
		t.Fatal("unknown Mac conditional concealed dependency")
	}
}
func TestExpressions(t *testing.T) {
	env := map[string]truth{"mac": yes, "vba7": yes, "win32": no}
	for s, want := range map[string]truth{"Mac And VBA7": yes, "Not Mac": no, "(Mac Or Win32) And VBA7": yes, "Mac = True": yes, "foo": unknown, "Win32 And foo": no, "Mac > 0": unknown} {
		if got := evaluate(s, env); got != want {
			t.Fatal(s, got, want)
		}
	}
}
