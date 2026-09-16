package native

import (
	"context"
	"fmt"
	"sort"
	"time"
)

// Probe launches app-owned hidden Word sessions back to back and reports
// whether every launch handshook, how long each took, and whether desktop
// isolation held. It is the gate for "WordUp can invoke Word headlessly on
// this machine": a failure is a product defect to diagnose, never a reason
// to report Word as unavailable. Each launch uses a fresh session and is
// closed before the next begins; nothing is left running.
func Probe(ctx context.Context, opt Options, count int) (map[string]any, error) {
	if count <= 0 {
		count = 1
	}
	if count > 50 {
		return nil, fmt.Errorf("probe count must be 1..50")
	}
	opt.Execute = false
	opt.Visible = false
	launches := make([]map[string]any, 0, count)
	startups := []float64{}
	faults := []any{}
	versions := map[string]bool{}
	isolated, closeErrors := 0, 0
	for i := 0; i < count; i++ {
		if ctx.Err() != nil {
			return nil, ctx.Err()
		}
		record := map[string]any{"index": i + 1}
		began := time.Now()
		h, err := Start(ctx, opt)
		if err != nil {
			f := fault(err)
			record["error"] = f
			record["elapsed_ms"] = ms(began)
			faults = append(faults, map[string]any{"index": i + 1, "error": f})
			launches = append(launches, record)
			continue
		}
		info := h.Info()
		if timing, ok := info["startup_timing_ms"].(map[string]float64); ok {
			record["startup_ms"] = timing
			startups = append(startups, timing["total"])
		} else {
			record["startup_ms"] = map[string]float64{"total": ms(began)}
			startups = append(startups, ms(began))
		}
		if v, ok := info["word_version"].(string); ok {
			versions[v] = true
			record["word_version"] = v
		}
		iso, _ := info["desktop_isolation"].(bool)
		record["desktop_isolation"] = iso
		if iso {
			isolated++
		}
		if degraded, ok := info["degraded_isolation"]; ok {
			record["degraded_isolation"] = degraded
			record["degraded_isolation_cause"] = info["degraded_isolation_cause"]
		}
		record["word_pid"] = info["pid"]
		closeBegan := time.Now()
		closeErr := h.Close()
		record["close_ms"] = ms(closeBegan)
		if closeErr != nil {
			f := fault(closeErr)
			record["close_error"] = f
			closeErrors++
			faults = append(faults, map[string]any{"index": i + 1, "close_error": f})
		}
		launches = append(launches, record)
	}
	sort.Float64s(startups)
	summary := map[string]any{
		"requested":            count,
		"launched":             len(startups),
		"failed":               count - len(startups),
		"close_errors":         closeErrors,
		"desktop_isolated":     isolated,
		"word_versions":        keys(versions),
		"startup_ms_min":       percentile(startups, 0),
		"startup_ms_p50":       percentile(startups, 50),
		"startup_ms_p95":       percentile(startups, 95),
		"startup_ms_max":       percentile(startups, 100),
		"launches":             launches,
		"available":            Available(),
		"headless_requirement": "every launch must handshake with owned hidden Word and close without leaving processes",
	}
	if len(faults) > 0 {
		summary["faults"] = faults
		return summary, Fail("native_launch_failed", fmt.Sprintf("%d of %d owned Word launches failed or did not clean up", len(faults), count), summary)
	}
	return summary, nil
}

func ms(t time.Time) float64 { return float64(time.Since(t).Microseconds()) / 1000 }

func percentile(sorted []float64, p int) float64 {
	if len(sorted) == 0 {
		return 0
	}
	if p <= 0 {
		return sorted[0]
	}
	if p >= 100 {
		return sorted[len(sorted)-1]
	}
	return sorted[(len(sorted)-1)*p/100]
}

func keys(m map[string]bool) []string {
	out := make([]string, 0, len(m))
	for k := range m {
		out = append(out, k)
	}
	sort.Strings(out)
	return out
}
