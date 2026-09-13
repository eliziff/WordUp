//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"strconv"
	"strings"
	"time"
)

// Inspect existing hyperlink-backed TOC results without updating their fields.
func (h *wordHost) inspectTOC(op Operation) (any, error) {
	if op.Target == "" || op.Target == "app" {
		return nil, fmt.Errorf("toc.inspect requires a document target")
	}
	d, err := h.object(op.Target)
	if err != nil {
		return nil, err
	}
	started := time.Now()
	paginationMS := float64(0)
	if op.Named["repaginate"] == true {
		v, e := d.call("Repaginate")
		v.clear()
		if e != nil {
			return nil, e
		}
		paginationMS = float64(time.Since(started).Microseconds()) / 1000
	}
	tables, err := objectProperty(d, "TablesOfContents")
	if err != nil {
		return nil, err
	}
	defer tables.release()
	count, err := scalarNumber(tables, "Count")
	if err != nil {
		return nil, err
	}
	bookmarks, err := objectProperty(d, "Bookmarks")
	if err != nil {
		return nil, err
	}
	defer bookmarks.release()
	rows := []any{}
	checked, mismatches, unresolved := 0, 0, 0
	result := map[string]any{"source": "native Word TOC hyperlinks and bookmark page information", "table_count": count, "repaginated": op.Named["repaginate"] == true, "fields_updated": false}
	for i := 1; i <= int(count); i++ {
		if len(rows) >= 2000 || time.Since(started) > 5*time.Second {
			err = fmt.Errorf("TOC inspection exceeded 2000 entries or five seconds")
			break
		}
		err = func() error {
			toc, e := objectProperty(tables, "Item", i)
			if e != nil {
				return e
			}
			defer toc.release()
			rangeObject, e := objectProperty(toc, "Range")
			if e != nil {
				return e
			}
			defer rangeObject.release()
			text, e := tocTextProperty(rangeObject, "Text")
			if e != nil {
				return e
			}
			if len(text) > 2<<20 {
				return fmt.Errorf("TOC text exceeds two MiB inspection budget")
			}
			unlinked := map[string]int{}
			paragraphs := strings.Split(text, "\r")
			entryCount := 0
			for _, paragraph := range paragraphs {
				paragraph = strings.TrimRight(paragraph, "\n\a")
				if strings.TrimSpace(paragraph) != "" {
					entryCount++
					if entryCount > 2000 {
						return fmt.Errorf("TOC exceeds 2000 paragraph inspection budget")
					}
					unlinked[paragraph]++
				}
			}
			links, e := objectProperty(rangeObject, "Hyperlinks")
			if e != nil {
				return e
			}
			defer links.release()
			n, e := scalarNumber(links, "Count")
			if e != nil {
				return e
			}
			if n == 0 {
				rows = append(rows, map[string]any{"table": i, "status": "unresolved", "reason": "TOC has no inspectable hyperlinks"})
				unresolved++
			}
			for j := 1; j <= int(n); j++ {
				if len(rows) >= 2000 || time.Since(started) > 5*time.Second {
					return Fail("toc_inspection_budget", "TOC inspection exceeded 2000 entries or five seconds", nil)
				}
				row, e := inspectTOCLink(links, bookmarks, j)
				row["table"], row["link"] = i, j
				rows = append(rows, row)
				if e != nil {
					return e
				}
				text, _ := row["text"].(string)
				text = strings.TrimRight(text, "\r\n\a")
				if unlinked[text] > 0 {
					unlinked[text]--
				}
				switch row["status"] {
				case "match":
					checked++
				case "mismatch":
					checked++
					mismatches++
				default:
					unresolved++
				}
			}
			for _, text := range paragraphs {
				text = strings.TrimRight(text, "\n\a")
				count := unlinked[text]
				if count > 0 {
					rows = append(rows, map[string]any{"table": i, "status": "unresolved", "reason": "TOC paragraph has no inspectable hyperlink", "text": text, "count": count})
					unresolved += count
					unlinked[text] = 0
				}
			}
			return nil
		}()
		if err != nil {
			break
		}
	}
	result["entries"], result["checked"], result["mismatches"], result["unresolved"] = rows, checked, mismatches, unresolved
	result["complete"] = err == nil && count > 0 && unresolved == 0
	result["timing_ms"] = map[string]any{"repaginate": paginationMS, "inspect": float64(time.Since(started).Microseconds())/1000 - paginationMS}
	if err != nil {
		return result, Fail("toc_inspection_failed", err.Error(), result)
	}
	return result, nil
}

func inspectTOCLink(links, bookmarks dispatch, index int) (map[string]any, error) {
	row := map[string]any{"status": "unresolved"}
	link, err := objectProperty(links, "Item", index)
	if err != nil {
		return row, err
	}
	defer link.release()
	name, err := tocTextProperty(link, "SubAddress")
	if err != nil {
		return row, err
	}
	row["bookmark"] = name
	r, err := objectProperty(link, "Range")
	if err != nil {
		return row, err
	}
	defer r.release()
	text, err := tocTextProperty(r, "Text")
	if err != nil {
		return row, err
	}
	row["text"] = text
	text = strings.TrimRight(text, "\r\n\a")
	_, page, found := strings.Cut(text, "\t")
	if last := strings.LastIndexByte(text, '\t'); last >= 0 {
		page = text[last+1:]
	}
	row["displayed_page"] = page
	wanted, err := strconv.Atoi(page)
	if !found || err != nil {
		row["reason"] = "No decimal page label; numbering format is not evaluated"
		return row, nil
	}
	v, err := bookmarks.call("Exists", name)
	if err != nil {
		return row, err
	}
	exists, err := v.value(0)
	v.clear()
	if err != nil {
		return row, err
	}
	if exists != true {
		row["reason"] = "Destination bookmark is missing"
		return row, nil
	}
	mark, err := objectProperty(bookmarks, "Item", name)
	if err != nil {
		return row, err
	}
	defer mark.release()
	destination, err := objectProperty(mark, "Range")
	if err != nil {
		return row, err
	}
	defer destination.release()
	v, err = destination.get("Information", 1) // wdActiveEndAdjustedPageNumber
	if err != nil {
		return row, err
	}
	actual, err := v.value(0)
	v.clear()
	if err != nil {
		return row, err
	}
	row["destination_page"] = actual
	row["status"] = "mismatch"
	if fmt.Sprint(actual) == strconv.Itoa(wanted) {
		row["status"] = "match"
	}
	return row, nil
}

func tocTextProperty(d dispatch, member string) (string, error) {
	v, err := d.get(member)
	if err != nil {
		return "", err
	}
	defer v.clear()
	value, err := v.value(0)
	if err != nil {
		return "", err
	}
	text, ok := value.(string)
	if !ok {
		return "", fmt.Errorf("%s did not return text", member)
	}
	return text, nil
}
