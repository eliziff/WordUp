package native

// Keep the runtime failure primary, but never turn failed cleanup into success.
func evaluationCleanupError(primary, cleanup error, path string) error {
	if cleanup == nil {
		return primary
	}
	details := map[string]any{"scratch_path": path, "cleanup_error": fault(cleanup)}
	if primary == nil {
		return Fail("scratch_cleanup_failed", "Scratch evaluation finished but its add-in could not be unloaded", details)
	}
	f := fault(primary)
	if original, ok := f.Details.(map[string]any); ok {
		for key, value := range original {
			details[key] = value
		}
	} else if f.Details != nil {
		details["runtime_details"] = f.Details
	}
	return Fail(f.Code, f.Message, details)
}
