package native

import "testing"

func TestProfilePreservesPartialFailureAndValidatesBeforeExecution(t *testing.T) {
	calls := 0
	call := func(op Operation) (any, error) {
		calls++
		if calls == 2 {
			return nil, Fail("seeded", "failure", nil)
		}
		return "first result", nil
	}
	_, err := profileOperations([]Operation{{Op: "get"}, {Op: "profile"}}, call)
	if err == nil || calls != 0 {
		t.Fatal("invalid sequence executed")
	}
	_, err = profileOperations([]Operation{{Op: "get"}, {Op: "run"}, {Op: "get"}}, call)
	f, ok := err.(*Fault)
	if !ok || calls != 2 {
		t.Fatalf("did not stop on failure: %v", err)
	}
	rows := f.Details.(map[string]any)["observations"].([]map[string]any)
	if len(rows) != 2 || rows[0]["result"] != "first result" || rows[1]["error"].(*Fault).Code != "seeded" {
		t.Fatalf("lost partial evidence: %v", rows)
	}
}
