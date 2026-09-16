package journal

import (
	"embed"
	"encoding/json"
	"fmt"
	"math"
	"sort"
	"strings"
)

// Inferred profiles are produced by tools/journal-profiles/infer_profiles.py
// from the published 2025/2026 articles of each olj-db journal. Every field
// carries the share of measured articles that agree; a null value means the
// evidence fell under the confidence floor and the generated macro must not
// guess. The files are bundled so journal.create needs neither the database
// nor the private article packages.
//
//go:embed profiles/*.json
var profileFS embed.FS

type measuredFloat struct {
	Value      *float64 `json:"value"`
	Confidence float64  `json:"confidence"`
	Articles   int      `json:"articles"`
}

type measuredString struct {
	Value      *string `json:"value"`
	Confidence float64 `json:"confidence"`
	Articles   int     `json:"articles"`
}

type measuredBool struct {
	Value      *bool   `json:"value"`
	Confidence float64 `json:"confidence"`
	Articles   int     `json:"articles"`
}

type measuredStrings struct {
	Value      []string `json:"value"`
	Confidence float64  `json:"confidence"`
	Articles   int      `json:"articles"`
}

type counted struct {
	Value      *string        `json:"value"`
	Counts     map[string]int `json:"counts"`
	Confidence float64        `json:"confidence"`
}

// Inferred mirrors profiles/<dataset>.json. Only the fields the generator
// consumes are decoded; evidence lists stay in the file for readers.
type Inferred struct {
	Dataset string `json:"dataset"`
	Journal string `json:"journal"`
	Abbrev  string `json:"abbrev"`
	Status  string `json:"status"`
	// Article ids are numbers in the database and strings in older runs.
	ArticlesMeasured []json.RawMessage `json:"articles_measured"`
	Page             struct {
		WidthPT    *float64 `json:"width_pt"`
		HeightPT   *float64 `json:"height_pt"`
		Confidence float64  `json:"confidence"`
		Articles   int      `json:"articles"`
	} `json:"page"`
	Margins struct {
		Mirrored  *bool         `json:"mirrored"`
		LeftOdd   measuredFloat `json:"left_odd_pages"`
		LeftEven  measuredFloat `json:"left_even_pages"`
		RightOdd  measuredFloat `json:"right_odd_pages"`
		RightEven measuredFloat `json:"right_even_pages"`
		Top       measuredFloat `json:"top"`
		Bottom    measuredFloat `json:"bottom"`
	} `json:"margins_in"`
	Body struct {
		Font         measuredString `json:"font"`
		SizePT       measuredFloat  `json:"size_pt"`
		LineHeightPT measuredFloat  `json:"line_height_pt"`
	} `json:"body"`
	Footnotes struct {
		Font                measuredString `json:"font"`
		SizePT              measuredFloat  `json:"size_pt"`
		Present             measuredBool   `json:"present"`
		NumberStyle         measuredString `json:"number_style"`
		IndentVsBodyIn      measuredFloat  `json:"indent_vs_body_in"`
		FirstPageAuthorNote measuredBool   `json:"first_page_author_note"`
	} `json:"footnotes"`
	Headings struct {
		NumberingLevels   measuredStrings `json:"numbering_levels"`
		TopLevelCase      measuredString  `json:"top_level_case"`
		TopLevelSizePT    measuredFloat   `json:"top_level_size_pt"`
		TopLevelAlignment measuredString  `json:"top_level_alignment"`
	} `json:"headings"`
	BlockQuotes struct {
		IndentIn measuredFloat `json:"indent_in"`
		SizePT   measuredFloat `json:"size_pt"`
	} `json:"block_quotes"`
	RunningHeads struct {
		OddPages   measuredStrings `json:"odd_pages"`
		EvenPages  measuredStrings `json:"even_pages"`
		PageNumber measuredString  `json:"page_number"`
	} `json:"running_heads"`
	TitleBlock struct {
		TitleSizePT     measuredFloat  `json:"title_size_pt"`
		TitleAlignment  measuredString `json:"title_alignment"`
		TitleCase       measuredString `json:"title_case"`
		AuthorCase      measuredString `json:"author_case"`
		AuthorAlignment measuredString `json:"author_alignment"`
	} `json:"title_block"`
	TableOfContents measuredBool       `json:"table_of_contents"`
	Citations       map[string]counted `json:"-"`
	RawCitations    json.RawMessage    `json:"citations"`
}

// Layout is the typesetting contract a generated Setup applies. Zero values
// mean "no evidence": the corresponding stage leaves the document alone and
// the field is listed in Profile.EvidenceGaps.
type Layout struct {
	PageWidthPT   float64 `json:"page_width_pt,omitempty"`
	PageHeightPT  float64 `json:"page_height_pt,omitempty"`
	MirrorMargins bool    `json:"mirror_margins,omitempty"`
	// Inside/Outside apply when MirrorMargins; Left/Right otherwise.
	TopIn        float64 `json:"top_in,omitempty"`
	BottomIn     float64 `json:"bottom_in,omitempty"`
	LeftIn       float64 `json:"left_in,omitempty"`
	RightIn      float64 `json:"right_in,omitempty"`
	InsideIn     float64 `json:"inside_in,omitempty"`
	OutsideIn    float64 `json:"outside_in,omitempty"`
	BodyLeadPT   float64 `json:"body_leading_pt,omitempty"`
	HeadingCase  string  `json:"heading_case,omitempty"`
	HeadingAlign string  `json:"heading_alignment,omitempty"`
	HeadingSize  float64 `json:"heading_size_pt,omitempty"`
	// HeadingNumbering lists the observed numbering patterns from the top
	// level down (roman_period, alpha_period, arabic_period, unnumbered...).
	HeadingNumbering    []string `json:"heading_numbering,omitempty"`
	QuoteIndentIn       float64  `json:"quote_indent_in,omitempty"`
	QuoteSizePT         float64  `json:"quote_size_pt,omitempty"`
	TitleSizePT         float64  `json:"title_size_pt,omitempty"`
	TitleAlign          string   `json:"title_alignment,omitempty"`
	TitleCase           string   `json:"title_case,omitempty"`
	AuthorAlign         string   `json:"author_alignment,omitempty"`
	AuthorCase          string   `json:"author_case,omitempty"`
	OddHead             []string `json:"odd_running_head,omitempty"`
	EvenHead            []string `json:"even_running_head,omitempty"`
	PageNumber          string   `json:"page_number,omitempty"`
	NoteNumberStyle     string   `json:"note_number_style,omitempty"`
	FirstPageAuthorNote bool     `json:"first_page_author_note,omitempty"`
	TableOfContents     bool     `json:"table_of_contents,omitempty"`
}

// Conventions are the citation habits counted in the journal's own
// footnotes. Each value is the winning form; Confidence is its share.
type Conventions struct {
	IbidCase            string             `json:"ibid_case,omitempty"`
	IbidPunctuation     string             `json:"ibid_punctuation,omitempty"`
	SupraForm           string             `json:"supra_form,omitempty"`
	ParagraphPinpoint   string             `json:"paragraph_pinpoint,omitempty"`
	PagePinpoint        string             `json:"page_pinpoint,omitempty"`
	RangeDash           string             `json:"range_dash,omitempty"`
	EmphasisNote        string             `json:"emphasis_note,omitempty"`
	SectionAbbreviation string             `json:"section_abbreviation,omitempty"`
	EtAl                string             `json:"et_al,omitempty"`
	Eg                  string             `json:"eg,omitempty"`
	Ie                  string             `json:"ie,omitempty"`
	SeeSignalCase       string             `json:"see_signal_case,omitempty"`
	OnlineAngleBrackets string             `json:"online_angle_brackets,omitempty"`
	Quotes              string             `json:"quotes,omitempty"`
	Confidence          map[string]float64 `json:"confidence,omitempty"`
}

// conventionFloor is the share of counted occurrences a form needs before a
// generated audit flags the other form; below it the journal is mixed and the
// rule stays off.
const conventionFloor = 0.8

// LoadInferred returns the bundled inferred profile for a dataset id.
func LoadInferred(id string) (Inferred, bool, error) {
	raw, err := profileFS.ReadFile("profiles/" + id + ".json")
	if err != nil {
		return Inferred{}, false, nil
	}
	var inf Inferred
	if err := json.Unmarshal(raw, &inf); err != nil {
		return Inferred{}, false, fmt.Errorf("profiles/%s.json: %w", id, err)
	}
	inf.Citations = map[string]counted{}
	if len(inf.RawCitations) > 0 {
		var fields map[string]json.RawMessage
		if err := json.Unmarshal(inf.RawCitations, &fields); err != nil {
			return Inferred{}, false, fmt.Errorf("profiles/%s.json citations: %w", id, err)
		}
		for key, value := range fields {
			var c counted
			if err := json.Unmarshal(value, &c); err != nil {
				continue // raw_counts and footnote_lines are not counted forms
			}
			inf.Citations[key] = c
		}
	}
	return inf, true, nil
}

// InferredIDs lists the datasets with a bundled inferred profile.
func InferredIDs() []string {
	entries, err := profileFS.ReadDir("profiles")
	if err != nil {
		return nil
	}
	var ids []string
	for _, e := range entries {
		name := e.Name()
		if strings.HasSuffix(name, ".json") && name != "index.json" {
			ids = append(ids, strings.TrimSuffix(name, ".json"))
		}
	}
	sort.Strings(ids)
	return ids
}

func floatOf(m measuredFloat) float64 {
	if m.Value == nil || math.IsNaN(*m.Value) || math.IsInf(*m.Value, 0) {
		return 0
	}
	return *m.Value
}

func stringOf(m measuredString) string {
	if m.Value == nil {
		return ""
	}
	return *m.Value
}

func boolOf(m measuredBool) bool {
	return m.Value != nil && *m.Value
}

// layoutFromInferred maps measurements to the typesetting contract, naming
// every field that had to stay empty.
func layoutFromInferred(inf Inferred) (Layout, []string) {
	var gaps []string
	var l Layout
	if inf.Page.WidthPT != nil && inf.Page.HeightPT != nil && inf.Page.Confidence >= 0.5 {
		l.PageWidthPT, l.PageHeightPT = *inf.Page.WidthPT, *inf.Page.HeightPT
	} else {
		gaps = append(gaps, "page size")
	}
	top, bottom := floatOf(inf.Margins.Top), floatOf(inf.Margins.Bottom)
	leftOdd, leftEven := floatOf(inf.Margins.LeftOdd), floatOf(inf.Margins.LeftEven)
	rightOdd, rightEven := floatOf(inf.Margins.RightOdd), floatOf(inf.Margins.RightEven)
	l.TopIn = top
	if top == 0 {
		gaps = append(gaps, "top margin")
	}
	l.BottomIn = bottom
	if bottom == 0 {
		gaps = append(gaps, "bottom margin")
	}
	mirrored := inf.Margins.Mirrored != nil && *inf.Margins.Mirrored
	switch {
	case mirrored && leftOdd > 0 && rightOdd > 0:
		// Odd (recto) pages carry the inside margin on the left.
		l.MirrorMargins = true
		l.InsideIn, l.OutsideIn = leftOdd, rightOdd
	case leftOdd > 0 || leftEven > 0:
		l.LeftIn = average(leftOdd, leftEven)
		l.RightIn = average(rightOdd, rightEven)
		if l.RightIn == 0 {
			gaps = append(gaps, "right margin")
		}
	default:
		gaps = append(gaps, "side margins")
	}
	l.BodyLeadPT = floatOf(inf.Body.LineHeightPT)
	if l.BodyLeadPT == 0 {
		gaps = append(gaps, "body leading")
	}
	l.HeadingCase = stringOf(inf.Headings.TopLevelCase)
	l.HeadingAlign = stringOf(inf.Headings.TopLevelAlignment)
	l.HeadingSize = floatOf(inf.Headings.TopLevelSizePT)
	l.HeadingNumbering = append([]string(nil), inf.Headings.NumberingLevels.Value...)
	if l.HeadingCase == "" {
		gaps = append(gaps, "heading case")
	}
	l.QuoteIndentIn = floatOf(inf.BlockQuotes.IndentIn)
	l.QuoteSizePT = floatOf(inf.BlockQuotes.SizePT)
	if l.QuoteIndentIn == 0 {
		gaps = append(gaps, "block quote indent")
	}
	l.TitleSizePT = floatOf(inf.TitleBlock.TitleSizePT)
	l.TitleAlign = stringOf(inf.TitleBlock.TitleAlignment)
	l.TitleCase = stringOf(inf.TitleBlock.TitleCase)
	l.AuthorAlign = stringOf(inf.TitleBlock.AuthorAlignment)
	l.AuthorCase = stringOf(inf.TitleBlock.AuthorCase)
	if l.TitleAlign == "" {
		gaps = append(gaps, "title block")
	}
	l.OddHead = append([]string(nil), inf.RunningHeads.OddPages.Value...)
	l.EvenHead = append([]string(nil), inf.RunningHeads.EvenPages.Value...)
	l.PageNumber = stringOf(inf.RunningHeads.PageNumber)
	if len(l.OddHead) == 0 && len(l.EvenHead) == 0 {
		gaps = append(gaps, "running heads")
	}
	if l.PageNumber == "" {
		gaps = append(gaps, "page number position")
	}
	l.NoteNumberStyle = stringOf(inf.Footnotes.NumberStyle)
	l.FirstPageAuthorNote = boolOf(inf.Footnotes.FirstPageAuthorNote)
	l.TableOfContents = boolOf(inf.TableOfContents)
	return l, gaps
}

func average(a, b float64) float64 {
	switch {
	case a > 0 && b > 0:
		return math.Round((a+b)/2*100) / 100
	case a > 0:
		return a
	default:
		return b
	}
}

// conventionsFromInferred keeps only forms that clearly dominate.
func conventionsFromInferred(inf Inferred) Conventions {
	c := Conventions{Confidence: map[string]float64{}}
	pick := func(key string) string {
		m, ok := inf.Citations[key]
		if !ok || m.Value == nil {
			return ""
		}
		total := 0
		for _, n := range m.Counts {
			total += n
		}
		if total < 10 || m.Confidence < conventionFloor {
			return ""
		}
		c.Confidence[key] = m.Confidence
		return *m.Value
	}
	c.IbidCase = pick("ibid_case")
	c.IbidPunctuation = pick("ibid_punctuation")
	c.SupraForm = pick("supra_form")
	c.ParagraphPinpoint = pick("paragraph_pinpoint")
	c.PagePinpoint = pick("page_pinpoint_label")
	c.RangeDash = pick("range_dash")
	c.EmphasisNote = pick("emphasis_note")
	c.SectionAbbreviation = pick("section_abbreviation")
	c.EtAl = pick("et_al")
	c.Eg = pick("eg")
	c.Ie = pick("ie")
	c.SeeSignalCase = pick("see_signal_case")
	c.OnlineAngleBrackets = pick("online_angle_brackets")
	c.Quotes = pick("quotes")
	return c
}

// WithInferred overlays the bundled inferred profile on a seed profile:
// fonts and sizes measured with high agreement replace the seed defaults, and
// the layout and citation conventions are attached for the generator.
func WithInferred(profile Profile) Profile {
	inf, ok, err := LoadInferred(profile.ID)
	if err != nil || !ok || inf.Status != "ok" {
		return profile
	}
	if font := stringOf(inf.Body.Font); font != "" && inf.Body.Font.Confidence >= 0.8 {
		profile.BodyFont = friendlyFont(font, profile.BodyFont)
	}
	if size := floatOf(inf.Body.SizePT); size > 0 && inf.Body.SizePT.Confidence >= 0.8 {
		profile.BodySizePT = size
	}
	if font := stringOf(inf.Footnotes.Font); font != "" && inf.Footnotes.Font.Confidence >= 0.8 {
		profile.NoteFont = friendlyFont(font, profile.NoteFont)
	}
	if size := floatOf(inf.Footnotes.SizePT); size > 0 && inf.Footnotes.SizePT.Confidence >= 0.8 {
		profile.NoteSizePT = size
	}
	layout, gaps := layoutFromInferred(inf)
	profile.Layout = &layout
	conventions := conventionsFromInferred(inf)
	profile.Conventions = &conventions
	profile.EvidenceGaps = gaps
	profile.EvidenceStatus = fmt.Sprintf("inferred from %d published 2025/2026 articles", len(inf.ArticlesMeasured))
	return profile
}
