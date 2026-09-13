package structure

// Adapted from legal-pdf-support pairing_support.rs::parse_heading_ladder,
// revision e1b060bdf9b92b9f6295bfc1e7933bd5a446af22 (MIT; PDF-LICENSE).
// A grammar assignment is evidence, not a decision that a paragraph is a heading.
type Interpretation struct {
	Family string `json:"family"`
	Value  int    `json:"value"`
}
type Assignment struct {
	Family string `json:"family"`
	Value  int    `json:"value"`
	Level  int    `json:"level"`
	Action string `json:"action"`
}

func HeadingLadder(candidates [][]Interpretation) []Assignment {
	stack := []Interpretation{}
	out := make([]Assignment, 0, len(candidates))
	for _, choices := range candidates {
		selected := Assignment{Action: "violation"}
		find := func(family string) int {
			for i := len(stack) - 1; i >= 0; i-- {
				if stack[i].Family == family {
					return i
				}
			}
			return -1
		}
		for phase := 0; phase < 4 && selected.Action == "violation"; phase++ {
			for _, c := range choices {
				if c.Family == "" || c.Value < 1 {
					continue
				}
				index := find(c.Family)
				action := ""
				switch phase {
				case 0:
					if index >= 0 && stack[index].Value+1 == c.Value {
						action = "increment"
					}
				case 1:
					if index < 0 && c.Value == 1 && len(stack) < 9 {
						action = "open_level"
					}
				case 2:
					if index >= 0 {
						if c.Value == 1 {
							action = "illegal_restart"
						} else if c.Value > stack[index].Value+1 {
							action = "jump_forward"
						}
					}
				case 3:
					if index < 0 && len(stack) < 9 {
						action = "open_midcounter"
					}
				}
				if action == "" {
					continue
				}
				if index < 0 {
					stack = append(stack, c)
					index = len(stack) - 1
				} else {
					stack = stack[:index+1]
					stack[index] = c
				}
				selected = Assignment{Family: c.Family, Value: c.Value, Level: index + 1, Action: action}
				break
			}
		}
		out = append(out, selected)
	}
	return out
}
