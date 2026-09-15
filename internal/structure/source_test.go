package structure

import (
	"strings"
	"testing"
)

func TestSourceCachesAlignedMainStoryTextWithTables(t *testing.T) {
	for _, marker := range []string{
		"WordUp structure contract 1.2.14.",
		"If document Is Nothing Then Err.Raise 91, \"WU_DetectStructure\", \"document is required\"",
		"main story is unavailable:",
		"main-story paragraphs are unavailable:",
		"Read the main story once.",
		"If Not storyTextAligned Then",
		"If hasTables Then",
		"relativeStart = startPosition - storyStart + 1",
		"If position <> storyEnd Then storyTextAligned = False",
	} {
		if !strings.Contains(Source, marker) {
			t.Fatalf("structure source lost marker %q", marker)
		}
	}
}
