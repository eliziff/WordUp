package native

import (
	"fmt"
	"time"
)

// profileOperations measures actual sequential execution, including any modal
// waits. It does not infer CPU time or certify the operations' results.
func profileOperations(steps []Operation, call func(Operation) (any, error)) (any, error) {
	if len(steps) == 0 || len(steps) > 1000 {
		return nil, fmt.Errorf("profile requires 1 to 1000 operations")
	}
	for _, step := range steps {
		if step.Op != "run" && step.Op != "eval" && step.Op != "get" && step.Op != "invoke" && step.Op != "put" {
			return nil, fmt.Errorf("profile does not support %q; use synchronous Word operations", step.Op)
		}
	}
	observations := []map[string]any{}
	started := time.Now()
	for i, step := range steps {
		t := time.Now()
		value, err := call(step)
		observations = append(observations, map[string]any{"index": i, "operation": step, "result": value, "error": fault(err), "duration_ms": float64(time.Since(t).Microseconds()) / 1000})
		if err != nil {
			return nil, Fail("profile_operation_failed", fmt.Sprintf("Profile operation %d failed: %v", i, err), map[string]any{"observations": observations, "cause": fault(err)})
		}
	}
	return map[string]any{"observations": observations, "duration_ms": float64(time.Since(started).Microseconds()) / 1000, "timing": "wall clock; includes modal waits"}, nil
}
