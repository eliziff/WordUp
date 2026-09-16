// Package journal provides the small, data-driven layer used to generate
// publication workspaces. The corpus is an input to discovery, never a
// runtime dependency of a finished DOTM.
package journal

import (
	"encoding/json"
	"fmt"
	"html"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

const CatalogSchema = 1

// VolumeEvidence is the volume/issue identity observed in the source tree.
// It deliberately contains no article text or private absolute paths.
type VolumeEvidence struct {
	Year     int    `json:"year"`
	Volume   string `json:"volume"`
	Issue    string `json:"issue"`
	Articles int    `json:"articles"`
}

type StyleEvidence struct {
	BodyFont              string  `json:"body_font,omitempty"`
	BodySizePT            float64 `json:"body_size_pt,omitempty"`
	NoteFont              string  `json:"note_font,omitempty"`
	NoteSizePT            float64 `json:"note_size_pt,omitempty"`
	Confidence            string  `json:"confidence"`
	ArticlesWithBodyStyle int     `json:"articles_with_body_style"`
	ArticlesWithNoteStyle int     `json:"articles_with_note_style"`
}

// ObservedEvidence is the conservative feature inventory used to decide
// which buttons should be visible by default. A false value means “not seen
// in this corpus slice”, not “the journal never uses it”.
type ObservedEvidence struct {
	Articles            int  `json:"articles"`
	HasFootnotes        bool `json:"has_footnotes"`
	HasBlockQuotes      bool `json:"has_block_quotes"`
	HasLists            bool `json:"has_lists"`
	HasContents         bool `json:"has_contents"`
	HasRunningFurniture bool `json:"has_running_furniture"`
	HasDropCaps         bool `json:"has_drop_caps"`
	HasTwoColumnPages   bool `json:"has_two_column_pages"`
}

// Profile is the public journal contract. It is small enough to edit by hand
// when an editor supplies a rule that the corpus cannot reveal, while the
// observed fields remain mechanically produced by Scan.
type Profile struct {
	ID              string           `json:"id"`
	Name            string           `json:"name"`
	Abbreviation    string           `json:"abbreviation,omitempty"`
	Scope           string           `json:"scope"`
	PermalinkPolicy string           `json:"permalink_policy"`
	Features        []string         `json:"features"`
	BodyFont        string           `json:"body_font"`
	BodySizePT      float64          `json:"body_size_pt"`
	NoteFont        string           `json:"note_font"`
	NoteSizePT      float64          `json:"note_size_pt"`
	CurrentYears    []int            `json:"current_years,omitempty"`
	Volumes         []VolumeEvidence `json:"volumes,omitempty"`
	StyleEvidence   StyleEvidence    `json:"style_evidence,omitempty"`
	Observed        ObservedEvidence `json:"observed,omitempty"`
	EvidenceStatus  string           `json:"evidence_status,omitempty"`
	// Layout and Conventions come from the bundled inferred profile (see
	// inferred.go); nil means no published-article evidence was available.
	Layout       *Layout      `json:"layout,omitempty"`
	Conventions  *Conventions `json:"conventions,omitempty"`
	EvidenceGaps []string     `json:"evidence_gaps,omitempty"`
}

type Catalog struct {
	Schema       int       `json:"schema"`
	Years        []int     `json:"years"`
	Source       string    `json:"source"`
	JournalCount int       `json:"journal_count"`
	ArticleCount int       `json:"article_count"`
	Journals     []Profile `json:"journals"`
}

var yearEvidence = regexp.MustCompile(`band_year[^0-9]+(20[0-9]{2})`)
var documentYearEvidence = regexp.MustCompile(`^\s*([0-9]{4})(?:\D|$)`)
var summaryDateEvidence = regexp.MustCompile(`"document_date_en"\s*:\s*"([^"]*)"`)

type fontRole struct {
	Active bool    `json:"active"`
	Font   string  `json:"font"`
	Size   float64 `json:"size"`
}

type summary struct {
	DocumentDate        string          `json:"document_date_en"`
	PageCount           int             `json:"page_count"`
	FontRoleBody        fontRole        `json:"font_role_body"`
	FontRoleNote        fontRole        `json:"font_role_note"`
	FontRoleCounts      map[string]int  `json:"font_role_counts"`
	BlockQuoteLineCount int             `json:"block_quote_line_count"`
	ListItemLineCount   int             `json:"list_item_line_count"`
	RunningFurniture    map[string]int  `json:"running_furniture_counts"`
	DropCapCount        int             `json:"drop_cap_merged_count"`
	TwoColumnFired      int             `json:"two_column_ladder_fired"`
	TOCOutline          json.RawMessage `json:"toc_outline"`
	FootnoteRecovery    struct {
		AnnotationCount int `json:"annotation_count"`
	} `json:"footnote_recovery"`
}

type styleAccumulator struct {
	bodyFonts    map[string]int
	noteFonts    map[string]int
	bodySizes    []float64
	noteSizes    []float64
	bodyArticles int
	noteArticles int
	observed     ObservedEvidence
}

func newAccumulator() styleAccumulator {
	return styleAccumulator{bodyFonts: map[string]int{}, noteFonts: map[string]int{}}
}

func (a *styleAccumulator) add(s summary) {
	a.observed.Articles++
	if s.FontRoleBody.Font != "" && s.FontRoleBody.Size > 0 {
		a.bodyFonts[s.FontRoleBody.Font]++
		a.bodySizes = append(a.bodySizes, s.FontRoleBody.Size)
		a.bodyArticles++
	}
	if s.FontRoleNote.Font != "" && s.FontRoleNote.Size > 0 {
		a.noteFonts[s.FontRoleNote.Font]++
		a.noteSizes = append(a.noteSizes, s.FontRoleNote.Size)
		a.noteArticles++
	}
	if s.FontRoleNote.Active || s.FontRoleCounts["note"] > 0 || s.FootnoteRecovery.AnnotationCount > 0 {
		a.observed.HasFootnotes = true
	}
	if s.BlockQuoteLineCount > 0 {
		a.observed.HasBlockQuotes = true
	}
	if s.ListItemLineCount > 0 {
		a.observed.HasLists = true
	}
	if len(s.TOCOutline) > 0 && string(s.TOCOutline) != "null" && string(s.TOCOutline) != "{}" {
		a.observed.HasContents = true
	}
	if len(s.RunningFurniture) > 0 {
		a.observed.HasRunningFurniture = true
	}
	if s.DropCapCount > 0 {
		a.observed.HasDropCaps = true
	}
	if s.TwoColumnFired > 0 {
		a.observed.HasTwoColumnPages = true
	}
}

func (a styleAccumulator) style() StyleEvidence {
	bodyFont := mostCommon(a.bodyFonts)
	noteFont := mostCommon(a.noteFonts)
	bodySize := median(a.bodySizes)
	noteSize := median(a.noteSizes)
	confidence := "none"
	if a.bodyArticles > 0 || a.noteArticles > 0 {
		confidence = "partial"
	}
	if a.bodyArticles > 0 && a.noteArticles > 0 {
		confidence = "observed"
	}
	return StyleEvidence{BodyFont: bodyFont, BodySizePT: bodySize, NoteFont: noteFont, NoteSizePT: noteSize, Confidence: confidence, ArticlesWithBodyStyle: a.bodyArticles, ArticlesWithNoteStyle: a.noteArticles}
}

func mostCommon(values map[string]int) string {
	best, bestCount := "", 0
	for value, count := range values {
		if count > bestCount || (count == bestCount && value < best) {
			best, bestCount = value, count
		}
	}
	return best
}

func median(values []float64) float64 {
	if len(values) == 0 {
		return 0
	}
	copyValues := append([]float64(nil), values...)
	sort.Float64s(copyValues)
	middle := len(copyValues) / 2
	if len(copyValues)%2 == 1 {
		return copyValues[middle]
	}
	return (copyValues[middle-1] + copyValues[middle]) / 2
}

func defaultFeatures(permalink string) []string {
	features := []string{"journal setup", "preflight", "style conversion", "heading and contents", "footnote and citation tools"}
	if permalink != "none" {
		features = append(features, "permalink assistant")
	}
	features = append(features, "supra tools", "tracked changes", "field refresh", "quality report")
	return features
}

type profileSeed struct {
	id, name, abbreviation, scope, permalink, bodyFont, noteFont string
	bodySize, noteSize                                           float64
}

// These names are stable IDs from the supplied corpus. Style values are
// conservative defaults from the current-year package summaries; Scan can
// replace them with fresh evidence without changing the generated engine.
var seeds = []profileSeed{
	{"ALTA-L-REV", "Alberta Law Review", "Alta L Rev", "law journal", "assistant", "Times New Roman", "Times New Roman", 10, 8},
	{"APPEAL", "Appeal: Review of Current Law and Law Reform", "Appeal", "law journal", "assistant", "Adobe Garamond Pro", "Myriad Pro", 10, 8},
	{"CAN-BAR-REV", "Canadian Bar Review", "Can Bar Rev", "law journal", "assistant", "Minion Pro", "Minion Pro", 11, 9},
	{"CAN-COMP-L-REV", "Canadian Competition Law Review", "Can Competition L Rev", "law journal", "assistant", "Minion Pro", "Minion Pro", 11, 9},
	{"CAN-J-FAM-L", "Canadian Journal of Family Law", "Can J Fam L", "law journal", "assistant", "Times New Roman", "Times New Roman", 12, 10},
	{"CAN-J-HUM-RTS", "Canadian Journal of Human Rights", "Can J Hum Rts", "law journal", "assistant", "Book Antiqua", "Book Antiqua", 11, 8},
	{"CAN-JL-JUR", "Canadian Journal of Law and Jurisprudence", "Can J L Jur", "law journal", "assistant", "Times New Roman", "Times New Roman", 10.5, 8.5},
	{"CAN-US-LJ", "Canada-United States Law Journal", "Can-USLJ", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"CIJS-ANN-REV", "Canadian Institute for the Study of Justice Annual Review", "CIJS Ann Rev", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 8},
	{"CJCA", "Canadian Journal of Commercial Arbitration", "CJCA", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"CJLS", "Canadian Journal of Law and Society", "CJLS", "law journal", "assistant", "Times New Roman", "Times New Roman", 10, 8},
	{"CONST-FORUM", "Constitutional Forum Constitutionnel", "Const Forum Const", "law journal", "assistant", "Minion Pro", "Minion Pro", 12, 10},
	{"DALHOUSIE-J-LEG-STUD", "Dalhousie Journal of Legal Studies", "Dal J Leg Stud", "law journal", "assistant", "Garamond", "Garamond", 11, 10},
	{"DALHOUSIE-LJ", "Dalhousie Law Journal", "Dal LJ", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 8},
	{"INDIGENOUS-LJ", "Indigenous Law Journal", "Indigenous LJ", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"JL-SOC-POLY", "Journal of Law and Social Policy", "JL Soc Poly", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"LAKEHEAD-LJ", "Lakehead Law Journal", "Lakehead LJ", "law journal", "assistant", "Sabon", "Sabon", 11, 9},
	{"MAN-LJ", "Manitoba Law Journal", "Man LJ", "law journal", "assistant", "Goudy Old Style", "Goudy Old Style", 11, 9},
	{"MCGILL-J-DISP-RES", "McGill Journal of Dispute Resolution", "McGill J Disp Res", "law journal", "assistant", "Times New Roman", "Times New Roman", 12, 10},
	{"MCGILL-J-SUSTAINABLE-DEV-L", "McGill Journal of Sustainable Development Law", "McGill J Sustainable Dev L", "law journal", "assistant", "Adobe Garamond Pro", "Adobe Garamond Pro", 11, 9},
	{"MCGILL-LJ", "McGill Law Journal", "McGill LJ", "law journal", "assistant", "Galliard", "Galliard", 11.5, 9},
	{"MCGILL-LJ-ERUDIT", "McGill Law Journal (Erudit archive)", "McGill LJ", "law journal", "assistant", "Galliard", "Galliard", 11, 9},
	{"MCGILL-LJ-HEALTH", "McGill Journal of Law and Health", "McGill JL & Health", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 10},
	{"OSGOODE-HALL-LJ", "Osgoode Hall Law Journal", "Osgoode Hall LJ", "law journal", "assistant", "Adobe Garamond Pro", "Adobe Garamond Pro", 10.5, 8.5},
	{"OTTAWA-L-REV", "Ottawa Law Review", "Ottawa L Rev", "law journal", "assistant", "FreightText Pro", "FreightText Pro", 11, 8.5},
	{"QUEENS-LJ", "Queen's Law Journal", "Queen's LJ", "law journal", "assistant", "Adobe Garamond Pro", "Adobe Garamond Pro", 10.5, 8.5},
	{"TRANSNATL-HUM-RTS-REV", "Transnational Human Rights Review", "Transnatl Hum Rts Rev", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"U-TORONTO-FAC-L-REV", "University of Toronto Faculty of Law Review", "UT Fac L Rev", "law journal", "assistant", "Plantin", "Plantin", 10, 8},
	{"U-TORONTO-J-L-EQUALITY", "University of Toronto Journal of Law and Equality", "UTJLE", "law journal", "assistant", "Garamond", "Garamond", 11, 9},
	{"UBC-L-REV", "UBC Law Review", "UBC L Rev", "law journal", "assistant", "Cambria", "Cambria", 11, 9},
	{"UNB-LJ", "University of New Brunswick Law Journal", "UNBLJ", "law journal", "assistant", "Times New Roman", "Times New Roman", 10, 8},
	{"W-J-LEG-STUD", "Windsor Journal of Legal Studies", "WJLS", "law journal", "assistant", "Book Antiqua", "Book Antiqua", 11, 9},
	{"WINDSOR-YB-ACCESS-JUST", "Windsor Yearbook of Access to Justice", "Windsor YB Access Just", "law journal", "assistant", "Times New Roman", "Times New Roman", 11, 9},
	{"WRONGFUL-CONVICTION-L-REV", "Wrongful Conviction Law Review", "Wrongful Conviction L Rev", "law journal", "assistant", "Times New Roman", "Times New Roman", 12, 11},
}

func seedProfile(s profileSeed) Profile {
	return Profile{ID: s.id, Name: s.name, Abbreviation: s.abbreviation, Scope: s.scope, PermalinkPolicy: s.permalink, Features: defaultFeatures(s.permalink), BodyFont: s.bodyFont, BodySizePT: s.bodySize, NoteFont: s.noteFont, NoteSizePT: s.noteSize, EvidenceStatus: "profile-default"}
}

func Profiles() []Profile {
	out := make([]Profile, 0, len(seeds))
	for _, seed := range seeds {
		out = append(out, WithInferred(seedProfile(seed)))
	}
	return out
}

func Get(id string) (Profile, error) {
	for _, profile := range Profiles() {
		if strings.EqualFold(profile.ID, id) {
			return profile, nil
		}
	}
	return Profile{}, fmt.Errorf("unknown journal %q", id)
}

func readSummary(path string) (summary, error) {
	file, err := os.Open(path)
	if err != nil {
		return summary{}, err
	}
	defer file.Close()
	var value summary
	decoder := json.NewDecoder(io.LimitReader(file, 8<<20))
	if err := decoder.Decode(&value); err != nil {
		return summary{}, err
	}
	return value, nil
}

// readSummaryDate is the cheap first pass used for the large corpus. The
// date field is emitted near the start of native summaries; old articles can
// therefore be rejected without decoding their complete, telemetry-heavy
// records. A summary whose field is not in this prefix falls back to the full
// decoder, preserving correctness for hand-authored inputs.
func readSummaryDate(path string) (string, bool, error) {
	file, err := os.Open(path)
	if err != nil {
		return "", false, err
	}
	defer file.Close()
	const prefixLimit = 16 << 10
	prefix, err := io.ReadAll(io.LimitReader(file, prefixLimit))
	if err != nil {
		return "", false, err
	}
	match := summaryDateEvidence.FindSubmatch(prefix)
	if len(match) != 2 {
		return "", false, nil
	}
	return string(match[1]), true, nil
}

func currentYear(provenance []byte, s summary, wanted map[int]bool) int {
	if y, ok := documentYear(s.DocumentDate); ok {
		if wanted[y] {
			return y
		}
		// An explicit, well-formed document date is authoritative even when it
		// is outside the requested window; do not let stale provenance move it
		// into a current volume.
		return 0
	}
	matches := yearEvidence.FindAllSubmatch(provenance, -1)
	for _, match := range matches {
		if y, err := strconv.Atoi(string(match[1])); err == nil && wanted[y] {
			return y
		}
	}
	return 0
}

func documentYear(value string) (int, bool) {
	value = strings.TrimSpace(value)
	if value == "" {
		return 0, false
	}
	if y, err := strconv.Atoi(value); err == nil && y >= 1900 && y <= 2200 {
		return y, true
	}
	match := documentYearEvidence.FindStringSubmatch(value)
	if len(match) != 2 {
		return 0, false
	}
	y, err := strconv.Atoi(match[1])
	return y, err == nil && y >= 1900 && y <= 2200
}

// Scan reads the final-contract storage layout and emits one compact profile
// per dataset with evidence for the requested years. It never reads article
// text; only the bounded provenance and native-summary records are consumed.
func Scan(root string, years []int) (Catalog, error) {
	if root == "" {
		return Catalog{}, fmt.Errorf("catalog path required")
	}
	absolute, err := filepath.Abs(root)
	if err != nil {
		return Catalog{}, err
	}
	info, err := os.Stat(absolute)
	if err != nil {
		return Catalog{}, err
	}
	if !info.IsDir() {
		return Catalog{}, fmt.Errorf("catalog path is not a directory: %s", root)
	}
	if len(years) == 0 {
		years = []int{2025, 2026}
	}
	wanted := map[int]bool{}
	for _, year := range years {
		if year < 1900 || year > 2200 {
			return Catalog{}, fmt.Errorf("invalid catalog year %d", year)
		}
		wanted[year] = true
	}
	years = append([]int(nil), years...)
	sort.Ints(years)
	type accumulator struct {
		profile  Profile
		styles   styleAccumulator
		volumes  map[string]*VolumeEvidence
		seenPath map[string]bool
	}
	groups := map[string]*accumulator{}
	articleCount := 0
	paths, err := filepath.Glob(filepath.Join(absolute, "*", "*", "*", "*", "provenance.json"))
	if err != nil {
		return Catalog{}, err
	}
	for _, path := range paths {
		rel, relErr := filepath.Rel(absolute, path)
		if relErr != nil {
			return Catalog{}, relErr
		}
		parts := strings.Split(filepath.ToSlash(rel), "/")
		if len(parts) < 5 {
			continue
		}
		dataset, volume, issue, article := parts[0], parts[1], parts[2], parts[3]
		summaryPath := filepath.Join(filepath.Dir(path), "digitalborn_native_summary.json")
		date, dateKnown, dateErr := readSummaryDate(summaryPath)
		if dateErr != nil && !os.IsNotExist(dateErr) {
			return Catalog{}, fmt.Errorf("%s summary date: %w", rel, dateErr)
		}
		if dateKnown {
			if year, known := documentYear(date); known && !wanted[year] {
				continue
			}
		}
		s, summaryErr := readSummary(summaryPath)
		if summaryErr != nil && !os.IsNotExist(summaryErr) {
			return Catalog{}, fmt.Errorf("%s summary: %w", rel, summaryErr)
		}
		year := currentYear(nil, s, wanted)
		// The summary carries the article date in a small, typed record. Only
		// fall back to the much larger provenance record when that field is
		// absent; this keeps discovery fast on the full image-rich corpus.
		if year == 0 {
			if _, known := documentYear(s.DocumentDate); known {
				continue
			}
			b, readErr := os.ReadFile(path)
			if readErr != nil {
				return Catalog{}, readErr
			}
			year = currentYear(b, s, wanted)
		}
		if year == 0 {
			continue
		}
		group := groups[dataset]
		if group == nil {
			profile, profileErr := Get(dataset)
			if profileErr != nil {
				profile = Profile{ID: dataset, Name: dataset, Scope: "unclassified", PermalinkPolicy: "assistant", Features: defaultFeatures("assistant"), BodyFont: "Times New Roman", BodySizePT: 11, NoteFont: "Times New Roman", NoteSizePT: 9, EvidenceStatus: "unrecognized-dataset"}
			}
			group = &accumulator{profile: profile, styles: newAccumulator(), volumes: map[string]*VolumeEvidence{}, seenPath: map[string]bool{}}
			groups[dataset] = group
		}
		if group.seenPath[rel] {
			continue
		}
		group.seenPath[rel] = true
		group.styles.add(s)
		group.profile.EvidenceStatus = "corpus-observed"
		key := fmt.Sprintf("%04d\x00%s\x00%s", year, volume, issue)
		v := group.volumes[key]
		if v == nil {
			v = &VolumeEvidence{Year: year, Volume: volume, Issue: issue}
			group.volumes[key] = v
		}
		v.Articles++
		articleCount++
		_ = article
	}
	out := Catalog{Schema: CatalogSchema, Years: years, Source: filepath.Clean(root)}
	for _, group := range groups {
		group.profile.Observed = group.styles.observed
		group.profile.StyleEvidence = group.styles.style()
		if group.profile.StyleEvidence.BodyFont != "" {
			group.profile.BodyFont = friendlyFont(group.profile.StyleEvidence.BodyFont, group.profile.BodyFont)
		}
		if group.profile.StyleEvidence.BodySizePT > 0 {
			group.profile.BodySizePT = group.profile.StyleEvidence.BodySizePT
		}
		if group.profile.StyleEvidence.NoteFont != "" {
			group.profile.NoteFont = friendlyFont(group.profile.StyleEvidence.NoteFont, group.profile.NoteFont)
		}
		if group.profile.StyleEvidence.NoteSizePT > 0 {
			group.profile.NoteSizePT = group.profile.StyleEvidence.NoteSizePT
		}
		for _, volume := range group.volumes {
			group.profile.Volumes = append(group.profile.Volumes, *volume)
			if !containsInt(group.profile.CurrentYears, volume.Year) {
				group.profile.CurrentYears = append(group.profile.CurrentYears, volume.Year)
			}
		}
		sort.Slice(group.profile.Volumes, func(i, j int) bool {
			a, b := group.profile.Volumes[i], group.profile.Volumes[j]
			if a.Year != b.Year {
				return a.Year < b.Year
			}
			if a.Volume != b.Volume {
				return a.Volume < b.Volume
			}
			return a.Issue < b.Issue
		})
		sort.Ints(group.profile.CurrentYears)
		out.Journals = append(out.Journals, group.profile)
	}
	sort.Slice(out.Journals, func(i, j int) bool { return out.Journals[i].ID < out.Journals[j].ID })
	out.JournalCount = len(out.Journals)
	out.ArticleCount = articleCount
	return out, nil
}

func containsInt(values []int, want int) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}

func friendlyFont(observed, fallback string) string {
	lower := strings.ToLower(observed)
	switch {
	case strings.Contains(lower, "timesnewroman"):
		return "Times New Roman"
	case strings.Contains(lower, "garamond"):
		return "Adobe Garamond Pro"
	case strings.Contains(lower, "minion"):
		return "Minion Pro"
	case strings.Contains(lower, "bookantiqua"):
		return "Book Antiqua"
	case strings.Contains(lower, "cambria"):
		return "Cambria"
	case strings.Contains(lower, "galliard"):
		return "Galliard"
	case strings.Contains(lower, "plantin"):
		return "Plantin"
	case strings.Contains(lower, "sabon"):
		return "Sabon"
	case strings.Contains(lower, "goudy"):
		return "Goudy Old Style"
	case strings.Contains(lower, "freight"):
		return "FreightText Pro"
	default:
		if fallback != "" {
			return fallback
		}
		return observed
	}
}

// EscapeXML is kept here so the macro generator and any future catalog UI
// agree on the same safe label encoding.
func EscapeXML(value string) string { return html.EscapeString(value) }
