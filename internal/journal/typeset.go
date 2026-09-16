package journal

import (
	"fmt"
	"strings"

	"github.com/eliziff/WordUp/internal/office"
)

// The typesetting module is generated from the journal's inferred Layout and
// Conventions. Every stage is a Public Sub the Typeset form runs inside one
// undo record and one batch; a stage whose evidence is missing reports that
// and changes nothing, so the checkbox stays honest.

func vbaBool(v bool) string {
	if v {
		return "True"
	}
	return "False"
}

func vbaSingle(v float64) string {
	return fmt.Sprintf("%.2f", v)
}

type typesetStage struct {
	Control string // checkbox name
	Caption string
	Macro   string
	Stage   string // batch stage label
	Enabled bool   // evidence available
}

func typesetStages(profile Profile) []typesetStage {
	l := profile.Layout
	c := profile.Conventions
	if l == nil {
		l = &Layout{}
	}
	if c == nil {
		c = &Conventions{}
	}
	return []typesetStage{
		{"chkPage", "Page size and margins", "WU_TypesetPageSetup", "Page size and margins", l.PageWidthPT > 0 || l.TopIn > 0 || l.LeftIn > 0 || l.InsideIn > 0},
		{"chkStyles", "House styles (body, headings, quotes, title)", "WU_TypesetStyles", "House styles", true},
		{"chkHeads", "Running heads and page numbers", "WU_TypesetRunningHeads", "Running heads", len(l.OddHead) > 0 || len(l.EvenHead) > 0 || l.PageNumber != ""},
		{"chkNoteTabs", "Tab after footnote numbers", "WU_TypesetFootnoteNumbers", "Footnote numbers", l.NoteNumberStyle == "number_gap" || l.NoteNumberStyle == "number_space"},
		{"chkQuotes", "Curly quotes", "WU_TypesetQuotes", "Curly quotes", c.Quotes == "curly"},
		{"chkAudit", "Audit citation conventions (highlight only)", "WU_TypesetCitationAudit", "Citation audit", conventionRules(*c) != ""},
		{"chkFix", "Fix unambiguous citation forms (Ibid case, et al, eg, ie)", "WU_TypesetCitationFix", "Citation fixes", c.IbidCase != "" || c.EtAl != "" || c.Eg != "" || c.Ie != ""},
	}
}

// columnSeparator separates the columns of a generated rule row. A visible
// token is used because the VBA editor may rewrite tab characters in source.
const columnSeparator = "||"

// conventionRules renders the audit table: each row is
// label || find text || wildcards || match case || whole word.
// The find text is the form the journal does NOT use.
func conventionRules(c Conventions) string {
	var rows []string
	add := func(label, find string, wildcards, matchCase, wholeWord bool) {
		rows = append(rows, strings.Join([]string{label, find, vbaBool(wildcards), vbaBool(matchCase), vbaBool(wholeWord)}, columnSeparator))
	}
	switch c.IbidCase {
	case "Ibid":
		add("lower-case ibid", "ibid", false, true, true)
	case "ibid":
		add("capitalised Ibid", "Ibid", false, true, true)
	}
	switch c.IbidPunctuation {
	case "Ibid.":
		add("Ibid followed by a comma", "Ibid,", false, true, false)
	case "Ibid,":
		add("Ibid followed by a period", "Ibid.", false, true, false)
	}
	switch c.SupraForm {
	case "supra note N":
		add("supra, note", "supra, note", false, false, false)
	case "supra, note N":
		add("supra note without comma", "supra note", false, false, false)
	}
	switch c.ParagraphPinpoint {
	case "at para N":
		add("at para. with period", "at para. ", false, false, false)
		add("at paras. with period", "at paras. ", false, false, false)
	case "at para. N":
		add("at para without period", "at para ", false, false, false)
	}
	switch c.PagePinpoint {
	case "at N":
		add("at p. page label", "at p[p.]{1,2} [0-9]", true, false, false)
	}
	switch c.RangeDash {
	case "en dash":
		add("hyphen in pinpoint range", "at [0-9]@-[0-9]", true, false, false)
	}
	switch c.EmphasisNote {
	case "[emphasis added]":
		add("(emphasis added) in parentheses", "(emphasis added)", false, false, false)
	case "(emphasis added)":
		add("[emphasis added] in brackets", "[emphasis added]", false, false, false)
	}
	switch c.SectionAbbreviation {
	case "s N":
		add("s. with period", "<[s]{1,2}. [0-9]", true, true, false)
	case "s. N":
		add("s without period", "<[s]{1,2} [0-9]", true, true, false)
	}
	switch c.EtAl {
	case "et al":
		add("et al. with period", "et al.", false, false, false)
	case "et al.":
		add("et al without period", "et al ", false, false, false)
	}
	switch c.Eg {
	case "eg":
		add("e.g. with periods", "e.g.", false, false, false)
	case "e.g.":
		add("eg without periods", "eg", false, true, true)
	}
	switch c.Ie {
	case "ie":
		add("i.e. with periods", "i.e.", false, false, false)
	case "i.e.":
		add("ie without periods", "ie", false, true, true)
	}
	switch c.OnlineAngleBrackets {
	case "online: <url>":
		add("online: without angle brackets", "online: http", false, false, false)
	}
	return strings.Join(rows, "\n")
}

// conventionFixes renders literal replacements that cannot change meaning:
// label || find || replace || match case || whole word.
func conventionFixes(c Conventions) string {
	var rows []string
	add := func(label, find, replace string, matchCase, wholeWord bool) {
		rows = append(rows, strings.Join([]string{label, find, replace, vbaBool(matchCase), vbaBool(wholeWord)}, columnSeparator))
	}
	switch c.IbidCase {
	case "Ibid":
		add("ibid to Ibid", "ibid", "Ibid", true, true)
	case "ibid":
		add("Ibid to ibid", "Ibid", "ibid", true, true)
	}
	switch c.EtAl {
	case "et al":
		add("et al. to et al", "et al.", "et al", false, false)
	case "et al.":
		add("et al to et al.", "et al ", "et al. ", false, false)
	}
	switch c.Eg {
	case "eg":
		add("e.g. to eg", "e.g.", "eg", false, false)
	case "e.g.":
		add("eg to e.g.", "eg", "e.g.", true, true)
	}
	switch c.Ie {
	case "ie":
		add("i.e. to ie", "i.e.", "ie", false, false)
	case "i.e.":
		add("ie to i.e.", "ie", "i.e.", true, true)
	}
	return strings.Join(rows, "\n")
}

func vbaTable(rows string) string {
	if rows == "" {
		return `""`
	}
	lines := strings.Split(rows, "\n")
	parts := make([]string, 0, len(lines))
	for _, line := range lines {
		parts = append(parts, vbaString(line))
	}
	return strings.Join(parts, " & vbLf & _\n    ")
}

func typesetSource(profile Profile) string {
	l := profile.Layout
	c := profile.Conventions
	if l == nil {
		l = &Layout{}
	}
	if c == nil {
		c = &Conventions{}
	}
	var b strings.Builder
	fmt.Fprintf(&b, `Attribute VB_Name = "WUJournalTypesetting"
Option Explicit
' Profile-driven typesetting for %s. Values were measured on the journal's
' published 2025/2026 articles (journal/profile.json lists the evidence). A
' zero or empty value means no evidence: that stage reports and changes
' nothing. Every stage is safe to run alone; the Typeset form runs the
' selected stages inside one undo record and one batch.

Public Const WU_TS_PAGE_WIDTH_PT As Single = %s
Public Const WU_TS_PAGE_HEIGHT_PT As Single = %s
Public Const WU_TS_MIRROR_MARGINS As Boolean = %s
Public Const WU_TS_MARGIN_TOP_IN As Single = %s
Public Const WU_TS_MARGIN_BOTTOM_IN As Single = %s
Public Const WU_TS_MARGIN_LEFT_IN As Single = %s
Public Const WU_TS_MARGIN_RIGHT_IN As Single = %s
Public Const WU_TS_MARGIN_INSIDE_IN As Single = %s
Public Const WU_TS_MARGIN_OUTSIDE_IN As Single = %s
Public Const WU_TS_BODY_LEADING_PT As Single = %s
Public Const WU_TS_HEADING_CASE As String = %s
Public Const WU_TS_HEADING_ALIGN As String = %s
Public Const WU_TS_HEADING_SIZE_PT As Single = %s
Public Const WU_TS_HEADING_NUMBERING As String = %s
Public Const WU_TS_QUOTE_INDENT_IN As Single = %s
Public Const WU_TS_QUOTE_SIZE_PT As Single = %s
Public Const WU_TS_TITLE_SIZE_PT As Single = %s
Public Const WU_TS_TITLE_ALIGN As String = %s
Public Const WU_TS_TITLE_CASE As String = %s
Public Const WU_TS_AUTHOR_ALIGN As String = %s
Public Const WU_TS_AUTHOR_CASE As String = %s
Public Const WU_TS_ODD_HEAD As String = %s
Public Const WU_TS_EVEN_HEAD As String = %s
Public Const WU_TS_PAGE_NUMBER As String = %s
Public Const WU_TS_NOTE_NUMBER_STYLE As String = %s
Public Const WU_TS_FIRST_PAGE_AUTHOR_NOTE As Boolean = %s
Public Const WU_TS_QUOTES As String = %s
Public Const WU_TS_HEAD_MARK As String = "WUJournalRunningHead"

' Audit rows: label, find text, wildcards, match case, whole word.
Private Const WU_TS_AUDIT_RULES As String = %s
' Fix rows: label, find text, replacement, match case, whole word.
Private Const WU_TS_FIX_RULES As String = %s

`, profile.Name,
		vbaSingle(l.PageWidthPT), vbaSingle(l.PageHeightPT), vbaBool(l.MirrorMargins),
		vbaSingle(l.TopIn), vbaSingle(l.BottomIn), vbaSingle(l.LeftIn), vbaSingle(l.RightIn), vbaSingle(l.InsideIn), vbaSingle(l.OutsideIn),
		vbaSingle(l.BodyLeadPT), vbaString(l.HeadingCase), vbaString(l.HeadingAlign), vbaSingle(l.HeadingSize), vbaString(strings.Join(l.HeadingNumbering, "|")),
		vbaSingle(l.QuoteIndentIn), vbaSingle(l.QuoteSizePT),
		vbaSingle(l.TitleSizePT), vbaString(l.TitleAlign), vbaString(l.TitleCase), vbaString(l.AuthorAlign), vbaString(l.AuthorCase),
		vbaString(strings.Join(l.OddHead, "|")), vbaString(strings.Join(l.EvenHead, "|")), vbaString(l.PageNumber),
		vbaString(l.NoteNumberStyle), vbaBool(l.FirstPageAuthorNote), vbaString(c.Quotes),
		vbaTable(conventionRules(*c)), vbaTable(conventionFixes(*c)))
	b.WriteString(typesetBody)
	return b.String()
}

// typesetBody is the profile-independent part of the module; it reads the
// constants above.
const typesetBody = `Public Sub WU_JournalTypesetFromRibbon(ByVal control As IRibbonControl)
    WU_JournalTypeset
End Sub

Public Sub WU_JournalTypeset()
    WUJournalTypeset.Show vbModeless
End Sub

' ---- page setup ------------------------------------------------------------

' Writes only values that differ: every PageSetup write relays out the section.
Public Sub WU_TypesetPageSetup()
    Dim doc As Document, section As Section, changed As Long
    Set doc = ActiveDocument
    If WU_TS_PAGE_WIDTH_PT <= 0 And WU_TS_MARGIN_TOP_IN <= 0 And WU_TS_MARGIN_LEFT_IN <= 0 And WU_TS_MARGIN_INSIDE_IN <= 0 Then
        WU_Notify "No page-size evidence for " & WU_JOURNAL_NAME & "; page setup left unchanged.", vbInformation
        Exit Sub
    End If
    For Each section In doc.Sections
        With section.PageSetup
            If WU_TS_PAGE_WIDTH_PT > 0 And WU_TS_PAGE_HEIGHT_PT > 0 Then
                If .Orientation <> wdOrientPortrait Then .Orientation = wdOrientPortrait: changed = changed + 1
                If Abs(.PageWidth - WU_TS_PAGE_WIDTH_PT) > 0.5 Then .PageWidth = WU_TS_PAGE_WIDTH_PT: changed = changed + 1
                If Abs(.PageHeight - WU_TS_PAGE_HEIGHT_PT) > 0.5 Then .PageHeight = WU_TS_PAGE_HEIGHT_PT: changed = changed + 1
            End If
            If WU_TS_MARGIN_TOP_IN > 0 Then If Abs(.TopMargin - InchesToPoints(WU_TS_MARGIN_TOP_IN)) > 0.5 Then .TopMargin = InchesToPoints(WU_TS_MARGIN_TOP_IN): changed = changed + 1
            If WU_TS_MARGIN_BOTTOM_IN > 0 Then If Abs(.BottomMargin - InchesToPoints(WU_TS_MARGIN_BOTTOM_IN)) > 0.5 Then .BottomMargin = InchesToPoints(WU_TS_MARGIN_BOTTOM_IN): changed = changed + 1
            If WU_TS_MIRROR_MARGINS Then
                If .MirrorMargins <> True Then .MirrorMargins = True: changed = changed + 1
                If WU_TS_MARGIN_INSIDE_IN > 0 Then If Abs(.LeftMargin - InchesToPoints(WU_TS_MARGIN_INSIDE_IN)) > 0.5 Then .LeftMargin = InchesToPoints(WU_TS_MARGIN_INSIDE_IN): changed = changed + 1
                If WU_TS_MARGIN_OUTSIDE_IN > 0 Then If Abs(.RightMargin - InchesToPoints(WU_TS_MARGIN_OUTSIDE_IN)) > 0.5 Then .RightMargin = InchesToPoints(WU_TS_MARGIN_OUTSIDE_IN): changed = changed + 1
            Else
                If WU_TS_MARGIN_LEFT_IN > 0 Then If Abs(.LeftMargin - InchesToPoints(WU_TS_MARGIN_LEFT_IN)) > 0.5 Then .LeftMargin = InchesToPoints(WU_TS_MARGIN_LEFT_IN): changed = changed + 1
                If WU_TS_MARGIN_RIGHT_IN > 0 Then If Abs(.RightMargin - InchesToPoints(WU_TS_MARGIN_RIGHT_IN)) > 0.5 Then .RightMargin = InchesToPoints(WU_TS_MARGIN_RIGHT_IN): changed = changed + 1
            End If
        End With
    Next section
    WU_Notify CStr(changed) & " page setup value(s) written.", vbInformation
End Sub

' ---- styles ----------------------------------------------------------------

' The shared style pass provisions and applies the named styles; this stage
' then shapes their definitions from the measured layout. Style definitions
' are document state, so each value is read before it is written.
Public Sub WU_TypesetStyles()
    Dim doc As Document
    Set doc = ActiveDocument
    WU_JournalApplyStyles
    WU_ShapeParagraphStyle doc, WU_JOURNAL_STYLE_BODY, WU_TS_BODY_LEADING_PT, "", "", 0, 0
    WU_ShapeParagraphStyle doc, WU_JOURNAL_STYLE_H1, 0, WU_TS_HEADING_CASE, WU_TS_HEADING_ALIGN, WU_TS_HEADING_SIZE_PT, 0
    WU_ShapeParagraphStyle doc, WU_JOURNAL_STYLE_QUOTATION, 0, "", "", WU_TS_QUOTE_SIZE_PT, WU_TS_QUOTE_INDENT_IN
    WU_ShapeParagraphStyle doc, WU_JOURNAL_STYLE_TITLE, 0, WU_TS_TITLE_CASE, WU_TS_TITLE_ALIGN, WU_TS_TITLE_SIZE_PT, 0
    WU_ShapeParagraphStyle doc, WU_JOURNAL_STYLE_AUTHOR, 0, WU_TS_AUTHOR_CASE, WU_TS_AUTHOR_ALIGN, 0, 0
    WU_Notify "House styles shaped from the " & WU_JOURNAL_NAME & " profile.", vbInformation
End Sub

Private Sub WU_ShapeParagraphStyle(ByVal doc As Document, ByVal styleName As String, ByVal leadingPT As Single, ByVal letterCase As String, ByVal align As String, ByVal sizePT As Single, ByVal indentIn As Single)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If st Is Nothing Then Exit Sub
    With st
        If leadingPT > 0 Then
            If .ParagraphFormat.LineSpacingRule <> wdLineSpaceExactly Or Abs(.ParagraphFormat.LineSpacing - leadingPT) > 0.05 Then
                .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                .ParagraphFormat.LineSpacing = leadingPT
            End If
        End If
        If sizePT > 0 Then If Abs(.Font.Size - sizePT) > 0.05 Then .Font.Size = sizePT
        If letterCase = "upper" Then
            If .Font.AllCaps <> True Then .Font.AllCaps = True
        ElseIf letterCase <> "" Then
            If .Font.AllCaps <> False Then .Font.AllCaps = False
        End If
        If align = "center" Then
            If .ParagraphFormat.Alignment <> wdAlignParagraphCenter Then .ParagraphFormat.Alignment = wdAlignParagraphCenter
        ElseIf align = "left" Then
            If .ParagraphFormat.Alignment <> wdAlignParagraphLeft Then .ParagraphFormat.Alignment = wdAlignParagraphLeft
        End If
        If indentIn > 0 Then
            If Abs(.ParagraphFormat.LeftIndent - InchesToPoints(indentIn)) > 0.5 Then .ParagraphFormat.LeftIndent = InchesToPoints(indentIn)
            If Abs(.ParagraphFormat.RightIndent - InchesToPoints(indentIn)) > 0.5 Then .ParagraphFormat.RightIndent = InchesToPoints(indentIn)
        End If
    End With
End Sub

' ---- running heads ---------------------------------------------------------

' Header content by page parity as the journal prints it. An existing header
' that this template did not write is left alone; ours carries a bookmark so
' a rerun replaces only its own text.
Public Sub WU_TypesetRunningHeads()
    Dim doc As Document, section As Section, written As Long, skipped As Long
    Dim oddText As String, evenText As String, useEven As Boolean
    Set doc = ActiveDocument
    If Len(WU_TS_ODD_HEAD) = 0 And Len(WU_TS_EVEN_HEAD) = 0 And Len(WU_TS_PAGE_NUMBER) = 0 Then
        WU_Notify "No running-head evidence for " & WU_JOURNAL_NAME & "; headers left unchanged.", vbInformation
        Exit Sub
    End If
    oddText = WU_HeadText(doc, WU_TS_ODD_HEAD)
    evenText = WU_HeadText(doc, WU_TS_EVEN_HEAD)
    useEven = (StrComp(oddText, evenText, vbBinaryCompare) <> 0)
    For Each section In doc.Sections
        If useEven Then If section.PageSetup.OddAndEvenPagesHeaderFooter <> True Then section.PageSetup.OddAndEvenPagesHeaderFooter = True
        If WU_WriteHead(section, wdHeaderFooterPrimary, oddText, True, WU_HeadMarkName(section.Index, "Primary")) Then written = written + 1 Else skipped = skipped + 1
        If useEven Then
            If WU_WriteHead(section, wdHeaderFooterEvenPages, evenText, False, WU_HeadMarkName(section.Index, "Even")) Then written = written + 1 Else skipped = skipped + 1
        End If
    Next section
    WU_Notify CStr(written) & " running head(s) written, " & CStr(skipped) & " left as the editor set them.", vbInformation
End Sub

' Compose header text from the measured content classes.
Private Function WU_HeadText(ByVal doc As Document, ByVal classes As String) As String
    Dim parts() As String, i As Long, piece As String, out As String
    If Len(classes) = 0 Then Exit Function
    parts = Split(classes, "|")
    For i = LBound(parts) To UBound(parts)
        piece = ""
        Select Case parts(i)
            Case "journal_name": piece = WU_JOURNAL_NAME
            Case "journal_abbrev": piece = WU_JOURNAL_NAME
            Case "volume_or_year": piece = WU_VolumeLabel(doc)
            Case "article_title": piece = WU_DocumentTitle(doc)
            Case "author": piece = WU_DocumentAuthor(doc)
        End Select
        If Len(piece) > 0 Then
            If Len(out) > 0 Then out = out & vbTab
            out = out & piece
        End If
    Next i
    WU_HeadText = out
End Function

Private Function WU_VolumeLabel(ByVal doc As Document) As String
    On Error Resume Next
    WU_VolumeLabel = doc.Variables("WUJournalVolumeLabel").Value
    If Len(WU_VolumeLabel) = 0 Then WU_VolumeLabel = "(" & Format$(Date, "yyyy") & ")"
    On Error GoTo 0
End Function

Private Function WU_DocumentTitle(ByVal doc As Document) As String
    Dim para As Paragraph
    On Error Resume Next
    WU_DocumentTitle = Trim$(doc.BuiltInDocumentProperties("Title").Value)
    On Error GoTo 0
    If Len(WU_DocumentTitle) > 0 Then Exit Function
    For Each para In doc.Paragraphs
        If StrComp(para.Style, WU_JOURNAL_STYLE_TITLE, vbTextCompare) = 0 Then
            WU_DocumentTitle = WU_ParagraphText(para)
            Exit Function
        End If
        If para.Range.Start > 4000 Then Exit For
    Next para
End Function

Private Function WU_DocumentAuthor(ByVal doc As Document) As String
    Dim para As Paragraph
    On Error Resume Next
    WU_DocumentAuthor = Trim$(doc.BuiltInDocumentProperties("Author").Value)
    On Error GoTo 0
    For Each para In doc.Paragraphs
        If StrComp(para.Style, WU_JOURNAL_STYLE_AUTHOR, vbTextCompare) = 0 Then
            WU_DocumentAuthor = WU_ParagraphText(para)
            Exit Function
        End If
        If para.Range.Start > 4000 Then Exit For
    Next para
End Function

Private Function WU_ParagraphText(ByVal para As Paragraph) As String
    Dim t As String
    t = para.Range.Text
    If Len(t) > 0 Then If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)
    WU_ParagraphText = Trim$(t)
End Function

' Bookmark names are unique per document, so each header and footer this
' template writes gets its own mark, all sharing the WU_TS_HEAD_MARK prefix.
Private Function WU_HeadMarkName(ByVal sectionIndex As Long, ByVal kind As String) As String
    WU_HeadMarkName = WU_TS_HEAD_MARK & "_S" & CStr(sectionIndex) & "_" & kind
End Function

' Whether a header or footer range carries one of this template's marks.
Public Function WU_HeadIsOwned(ByVal rng As Range) As Boolean
    Dim mark As Bookmark
    For Each mark In rng.Bookmarks
        If Left$(mark.Name, Len(WU_TS_HEAD_MARK)) = WU_TS_HEAD_MARK Then WU_HeadIsOwned = True: Exit Function
    Next mark
End Function

' Returns True when the header was written; False when an editor's header
' was left in place.
Private Function WU_WriteHead(ByVal section As Section, ByVal kind As Long, ByVal text As String, ByVal oddPage As Boolean, ByVal markName As String) As Boolean
    Dim header As HeaderFooter, rng As Range, existing As String, numberPosition As String, ours As Boolean
    Set header = section.Headers(kind)
    Set rng = header.Range
    existing = rng.Text
    If Len(existing) > 0 Then If Right$(existing, 1) = vbCr Then existing = Left$(existing, Len(existing) - 1)
    ours = WU_HeadIsOwned(rng)
    If Len(Trim$(existing)) > 0 And Not ours Then Exit Function
    numberPosition = WU_TS_PAGE_NUMBER
    rng.Text = text
    If Left$(numberPosition, 3) = "top" Then
        Select Case Mid$(numberPosition, 5)
            Case "center"
                rng.InsertAfter vbTab
                rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                WU_InsertPageField header, True
                rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Case "outer"
                If oddPage Then WU_AppendNumber section, header Else WU_PrependNumber header
            Case "inner"
                If oddPage Then WU_PrependNumber header Else WU_AppendNumber section, header
            Case Else
                WU_AppendNumber section, header
        End Select
    ElseIf Left$(numberPosition, 6) = "bottom" Then
        WU_WriteFooterNumber section, kind, numberPosition, oddPage, markName & "_Footer"
    End If
    header.Range.Bookmarks.Add markName, header.Range
    WU_WriteHead = True
End Function

Private Sub WU_InsertPageField(ByVal header As HeaderFooter, ByVal atEnd As Boolean)
    Dim rng As Range
    Set rng = header.Range
    rng.Collapse IIf(atEnd, wdCollapseEnd, wdCollapseStart)
    If atEnd Then rng.Move wdCharacter, -1
    header.Range.Fields.Add rng, wdFieldPage, , False
End Sub

Private Sub WU_AppendNumber(ByVal section As Section, ByVal header As HeaderFooter)
    Dim rng As Range
    Set rng = header.Range
    rng.InsertAfter vbTab
    ' A right tab at the text width keeps the number flush with the outer
    ' margin whatever the page size.
    rng.ParagraphFormat.TabStops.ClearAll
    rng.ParagraphFormat.TabStops.Add section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin, wdAlignTabRight
    WU_InsertPageField header, True
End Sub

Private Sub WU_PrependNumber(ByVal header As HeaderFooter)
    Dim rng As Range
    Set rng = header.Range
    rng.InsertBefore vbTab
    rng.Collapse wdCollapseStart
    header.Range.Fields.Add rng, wdFieldPage, , False
End Sub

Private Sub WU_WriteFooterNumber(ByVal section As Section, ByVal kind As Long, ByVal position As String, ByVal oddPage As Boolean, ByVal markName As String)
    Dim footer As HeaderFooter, rng As Range, existing As String
    Set footer = section.Footers(kind)
    Set rng = footer.Range
    existing = rng.Text
    If Len(existing) > 0 Then If Right$(existing, 1) = vbCr Then existing = Left$(existing, Len(existing) - 1)
    If Len(Trim$(existing)) > 0 And Not WU_HeadIsOwned(rng) Then Exit Sub
    rng.Text = ""
    Select Case Mid$(position, 8)
        Case "center": rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Case "outer": rng.ParagraphFormat.Alignment = IIf(oddPage, wdAlignParagraphRight, wdAlignParagraphLeft)
        Case "inner": rng.ParagraphFormat.Alignment = IIf(oddPage, wdAlignParagraphLeft, wdAlignParagraphRight)
        Case Else: rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End Select
    rng.Collapse wdCollapseStart
    footer.Range.Fields.Add rng, wdFieldPage, , False
    footer.Range.Bookmarks.Add markName, footer.Range
End Sub

' ---- footnote numbers ------------------------------------------------------

' The journal separates the note number from its text with a gap; in Word
' that is a tab after the reference mark. Notes already starting with a tab
' are untouched.
Public Sub WU_TypesetFootnoteNumbers()
    Dim doc As Document, note As Footnote, rng As Range, added As Long, firstChar As String
    Set doc = ActiveDocument
    If WU_TS_NOTE_NUMBER_STYLE <> "number_gap" And WU_TS_NOTE_NUMBER_STYLE <> "number_space" Then
        WU_Notify "No footnote-number evidence for " & WU_JOURNAL_NAME & "; notes left unchanged.", vbInformation
        Exit Sub
    End If
    For Each note In doc.Footnotes
        Set rng = note.Range
        If rng.Start < rng.End Then
            firstChar = Left$(rng.Text, 1)
            If firstChar <> vbTab Then
                If firstChar = " " Then
                    rng.Characters(1).Text = vbTab
                Else
                    rng.InsertBefore vbTab
                End If
                added = added + 1
            End If
        End If
    Next note
    WU_Notify CStr(added) & " footnote tab(s) added.", vbInformation
End Sub

' ---- quotes ----------------------------------------------------------------

' Straight quotes become the journal's curly quotes in every story that
' holds one; orientation follows the preceding character exactly as Word's
' own smart quotes do.
Public Sub WU_TypesetQuotes()
    Dim story As Range, storyText As String, count As Long
    If WU_TS_QUOTES <> "curly" Then
        WU_Notify "The " & WU_JOURNAL_NAME & " profile shows no quote preference; quotes left unchanged.", vbInformation
        Exit Sub
    End If
    For Each story In WU_Stories(ActiveDocument, "all")
        storyText = story.Text
        If InStr(1, storyText, """", vbBinaryCompare) > 0 Or InStr(1, storyText, "'", vbBinaryCompare) > 0 Then
            count = count + WU_CurlQuotes(story, """", ChrW(8220), ChrW(8221))
            count = count + WU_CurlQuotes(story, "'", ChrW(8216), ChrW(8217))
        End If
    Next story
    WU_Notify CStr(count) & " quote pass(es) applied.", vbInformation
End Sub

Private Function WU_CurlQuotes(ByVal story As Range, ByVal straight As String, ByVal opening As String, ByVal closing As String) As Long
    Dim probe As Range
    Set probe = story.Duplicate
    With probe.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = straight
        .Replacement.Text = opening
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        If .Execute(Replace:=wdReplaceAll) Then WU_CurlQuotes = WU_CurlQuotes + 1
        ' Word rejects a closing parenthesis inside a wildcard class and reads
        ' an unescaped exclamation mark there as negation, so the bracket case
        ' is a separate literal pass and the exclamation mark is escaped.
        .Text = "([A-Za-z0-9.,;:?\!])" & opening
        .Replacement.Text = "\1" & closing
        If .Execute(Replace:=wdReplaceAll) Then WU_CurlQuotes = WU_CurlQuotes + 1
        .MatchWildcards = False
        .Text = ")" & opening
        .Replacement.Text = ")" & closing
        If .Execute(Replace:=wdReplaceAll) Then WU_CurlQuotes = WU_CurlQuotes + 1
    End With
End Function

' ---- citation conventions --------------------------------------------------

' Highlights every occurrence of a form the journal does not use, in the
' footnote stories, and reports the count per rule. Nothing is rewritten.
Public Sub WU_TypesetCitationAudit()
    Dim rows() As String, cols() As String, i As Long, story As Range, hits As Long, total As Long, report As String
    If Len(WU_TS_AUDIT_RULES) = 0 Then
        WU_Notify "No citation conventions were measured with enough agreement for " & WU_JOURNAL_NAME & ".", vbInformation
        Exit Sub
    End If
    rows = Split(WU_TS_AUDIT_RULES, vbLf)
    For i = LBound(rows) To UBound(rows)
        cols = Split(rows(i), "||")
        hits = 0
        For Each story In WU_Stories(ActiveDocument, "notes")
            hits = hits + WU_HighlightMatches(story, cols(1), cols(2) = "True", cols(3) = "True", cols(4) = "True")
        Next story
        If hits > 0 Then report = report & cols(0) & ": " & CStr(hits) & vbCrLf
        total = total + hits
    Next i
    If total = 0 Then
        WU_Notify "Citation audit: every measured convention is followed.", vbInformation
    Else
        WU_Notify "Citation audit highlighted " & CStr(total) & " item(s):" & vbCrLf & report, vbInformation
    End If
End Sub

Private Function WU_HighlightMatches(ByVal story As Range, ByVal findText As String, ByVal wildcards As Boolean, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Long
    Dim probe As Range, limit As Long
    Set probe = story.Duplicate
    limit = story.End
    With probe.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .MatchWildcards = wildcards
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
    Do While probe.Find.Execute
        If probe.End > limit Or probe.End <= probe.Start Then Exit Do
        probe.HighlightColorIndex = wdYellow
        WU_HighlightMatches = WU_HighlightMatches + 1
        If probe.End >= limit Then Exit Do
        probe.Start = probe.End
        probe.End = limit
    Loop
End Function

' Number of audit rules and the fix targets, so a proof or a form can read
' what this profile enables without parsing the tables.
Public Function WU_TypesetAuditRuleCount() As Long
    If Len(WU_TS_AUDIT_RULES) = 0 Then Exit Function
    WU_TypesetAuditRuleCount = UBound(Split(WU_TS_AUDIT_RULES, vbLf)) + 1
End Function

Public Function WU_TypesetFixTargets() As String
    Dim rows() As String, cols() As String, i As Long
    If Len(WU_TS_FIX_RULES) = 0 Then Exit Function
    rows = Split(WU_TS_FIX_RULES, vbLf)
    For i = LBound(rows) To UBound(rows)
        cols = Split(rows(i), "||")
        WU_TypesetFixTargets = WU_TypesetFixTargets & "|" & cols(1) & ">" & cols(2)
    Next i
End Function

' Applies the literal replacements whose meaning cannot change (case of Ibid,
' periods in et al, eg and ie) in the footnote stories.
Public Sub WU_TypesetCitationFix()
    Dim rows() As String, cols() As String, i As Long, changed As Long, rules As Variant, n As Long
    If Len(WU_TS_FIX_RULES) = 0 Then
        WU_Notify "No unambiguous citation fixes were measured for " & WU_JOURNAL_NAME & ".", vbInformation
        Exit Sub
    End If
    rows = Split(WU_TS_FIX_RULES, vbLf)
    For i = LBound(rows) To UBound(rows)
        cols = Split(rows(i), "||")
        ReDim rules(0 To 0, 0 To 1)
        rules(0, 0) = cols(1): rules(0, 1) = cols(2)
        n = WU_ReplaceBatch(ActiveDocument, rules, "notes", cols(3) = "True", cols(4) = "True", False)
        changed = changed + n
    Next i
    WU_Notify CStr(changed) & " citation rule(s) applied.", vbInformation
End Sub
`

// typesetFormDesign is the checkbox task list, one box per stage, with a
// volume label for running heads and the ALR-style select-all/run pair.
func typesetFormDesign(profile Profile) office.Design {
	stages := typesetStages(profile)
	controls := []office.ControlDesign{
		{Name: "lblTitle", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 10.0, "Width": 400.0, "Height": 20.0, "Caption": "Typeset for " + profile.Name}},
		{Name: "lblNote", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 30.0, "Width": 400.0, "Height": 30.0, "Caption": "Measured on the journal's published articles. Unticked or greyed stages have no evidence and change nothing."}},
		{Name: "lblVolume", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 66.0, "Width": 150.0, "Height": 18.0, "Caption": "Volume label for running heads"}},
		{Name: "txtVolume", Type: "TextBox", Properties: map[string]any{"Left": 170.0, "Top": 64.0, "Width": 240.0, "Height": 20.0}},
	}
	top := 96.0
	for _, s := range stages {
		controls = append(controls, office.ControlDesign{Name: s.Control, Type: "CheckBox", Properties: map[string]any{
			"Left": 14.0, "Top": top, "Width": 396.0, "Height": 20.0, "Caption": s.Caption, "Value": vbaBool(s.Enabled), "Enabled": s.Enabled,
		}})
		top += 24
	}
	controls = append(controls,
		office.ControlDesign{Name: "cmdSelectAll", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": top + 8, "Width": 120.0, "Height": 26.0, "Caption": "Select all"}},
		office.ControlDesign{Name: "cmdRun", Type: "CommandButton", Properties: map[string]any{"Left": 144.0, "Top": top + 8, "Width": 150.0, "Height": 26.0, "Caption": "Run selected"}},
		office.ControlDesign{Name: "cmdClose", Type: "CommandButton", Properties: map[string]any{"Left": 304.0, "Top": top + 8, "Width": 106.0, "Height": 26.0, "Caption": "Close"}},
	)
	return office.Design{
		Name: "WUJournalTypeset",
		Mode: "replace",
		Properties: map[string]any{
			"Caption": "Typeset - " + profile.Name,
			"Width":   430.0,
			"Height":  top + 70,
		},
		Controls: controls,
	}
}

func typesetFormSource(profile Profile) string {
	stages := typesetStages(profile)
	var calls strings.Builder
	for _, s := range stages {
		fmt.Fprintf(&calls, "    If %s.Value Then\n        WU_BatchStage %s\n        %s\n    End If\n", s.Control, vbaString(s.Stage), s.Macro)
	}
	var selectAll strings.Builder
	for _, s := range stages {
		fmt.Fprintf(&selectAll, "    If %s.Enabled Then %s.Value = True\n", s.Control, s.Control)
	}
	return fmt.Sprintf(`Attribute VB_Name = "WUJournalTypeset"
Option Explicit

Private Sub cmdSelectAll_Click()
%sEnd Sub

Private Sub cmdRun_Click()
    Dim total As Long, ctrl As Control, report As Variant
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then If ctrl.Value Then total = total + 1
    Next ctrl
    If total = 0 Then
        MsgBox "Select at least one stage.", vbExclamation, "Typeset"
        Exit Sub
    End If
    On Error GoTo Failed
    If Len(Trim$(txtVolume.Text)) > 0 Then ActiveDocument.Variables("WUJournalVolumeLabel").Value = Trim$(txtVolume.Text)
    WU_BeginSafeEdit updating, opened, captured, "Typeset " & WU_JOURNAL_NAME
    WU_BatchStart total
%s
CleanUp:
    On Error Resume Next
    WU_BatchEnd
    WU_EndSafeEdit updating, opened, captured
    On Error GoTo 0
    report = WU_BatchReport()
    If failure <> 0 Then
        MsgBox failureText, vbExclamation, "Typeset"
    ElseIf report(0) Then
        MsgBox "Typesetting finished with warnings:" & vbCrLf & report(1), vbExclamation, "Typeset"
    Else
        MsgBox "Typesetting finished. Use one Undo to reverse every change.", vbInformation, "Typeset"
    End If
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
`, selectAll.String(), calls.String())
}
