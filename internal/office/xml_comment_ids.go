package office

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
)

const word14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const word15 = "http://schemas.microsoft.com/office/word/2012/wordml"
const wordCID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
const wordCEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"

// Generated IDs are mapped through comment identity, never discarded. Reply
// parents and durable-ID references must resolve to a declared comment.
type commentIDs struct {
	paragraphs map[string]string
	last       map[string]string
	durable    map[string]string
}

func xmlAttr(el xml.StartElement, space, local string) string {
	for _, a := range el.Attr {
		if a.Name == (xml.Name{Space: space, Local: local}) {
			return a.Value
		}
	}
	return ""
}

func readCommentIDs(raw []byte) (*commentIDs, error) {
	out := &commentIDs{map[string]string{}, map[string]string{}, map[string]string{}}
	d := xml.NewDecoder(bytes.NewReader(raw))
	comments := map[string]bool{}
	var records []xml.StartElement
	comment, last := "", ""
	paragraph := 0
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch el := token.(type) {
		case xml.StartElement:
			switch el.Name {
			case xml.Name{Space: wordXMLNamespace, Local: "comment"}:
				comment = xmlAttr(el, wordXMLNamespace, "id")
				if comment == "" || comments[comment] {
					return nil, fmt.Errorf("missing or duplicate comment identity %q", comment)
				}
				comments[comment], paragraph, last = true, 0, ""
			case xml.Name{Space: wordXMLNamespace, Local: "p"}:
				if comment != "" {
					paragraph++
					last = xmlAttr(el, word14, "paraId")
					if last != "" {
						if _, exists := out.paragraphs[last]; exists {
							return nil, fmt.Errorf("duplicate comment paragraph ID %q", last)
						}
						out.paragraphs[last] = fmt.Sprintf("\x00WordUpComment%sParagraph%d", comment, paragraph)
					}
				}
			case xml.Name{Space: wordCID, Local: "commentId"}:
				records = append(records, el.Copy())
			}
		case xml.EndElement:
			if el.Name == (xml.Name{Space: wordXMLNamespace, Local: "comment"}) {
				if last != "" {
					out.last[last] = out.paragraphs[last]
				}
				comment = ""
			}
		}
	}
	for _, el := range records {
		para, durable := xmlAttr(el, wordCID, "paraId"), xmlAttr(el, wordCID, "durableId")
		canonical, ok := out.last[para]
		if !ok || durable == "" || out.durable[durable] != "" {
			return nil, fmt.Errorf("invalid comment identity link: paragraph %q, durable ID %q", para, durable)
		}
		out.durable[durable] = canonical
	}
	return out, nil
}

func (ids *commentIDs) canonical(el xml.Name, at xml.Attr) (string, error) {
	var values map[string]string
	switch {
	case el == (xml.Name{Space: wordXMLNamespace, Local: "p"}) && at.Name == (xml.Name{Space: word14, Local: "paraId"}):
		if value, ok := ids.paragraphs[at.Value]; ok {
			return value, nil
		}
		return at.Value, nil
	case el == (xml.Name{Space: word15, Local: "commentEx"}) && at.Name.Space == word15 && (at.Name.Local == "paraId" || at.Name.Local == "paraIdParent"):
		values = ids.last
	case el == (xml.Name{Space: wordCID, Local: "commentId"}) && at.Name == (xml.Name{Space: wordCID, Local: "paraId"}):
		values = ids.last
	case el == (xml.Name{Space: wordCID, Local: "commentId"}) && at.Name == (xml.Name{Space: wordCID, Local: "durableId"}),
		el == (xml.Name{Space: wordCEX, Local: "commentExtensible"}) && at.Name == (xml.Name{Space: wordCEX, Local: "durableId"}):
		values = ids.durable
	default:
		return at.Value, nil
	}
	value, ok := values[at.Value]
	if !ok {
		return "", fmt.Errorf("unresolved comment link {%s}%s=%q", at.Name.Space, at.Name.Local, at.Value)
	}
	return value, nil
}
