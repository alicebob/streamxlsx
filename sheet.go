package streamxlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

type rowXML struct {
	Cells []Cell `xml:"c"`
	Index int    `xml:"r,attr"`
}

// These can be passed to `WriteRow()` if you want total control. WriteRow() will fill int the `.Ref` value.
// Exactly one of Value or InlineString should be set.
type Cell struct {
	Ref          string  `xml:"r,attr"` // "A1" &c.
	Type         string  `xml:"t,attr,omitempty"`
	Style        *int    `xml:"s,attr,omitempty"`
	Value        string  `xml:"v,omitempty"`
	InlineString *string `xml:"is>t,omitempty"`
	hyperlink    *hyperlink
}

func (c Cell) String() string {
	if c.InlineString != nil {
		return *c.InlineString
	}
	return c.Value
}

type hyperlink struct {
	RelID   string `xml:"r:id,attr"`
	Ref     string `xml:"ref,attr"`
	Display string `xml:"display,attr"`
	Tooltip string `xml:"tooltip,attr"`
	url     string
}

type sheetEncoder struct {
	enc        *xml.Encoder
	rows       int
	hyperlinks []hyperlink
	relations  []relationship
}

func newSheetEncoder(fh io.Writer) *sheetEncoder {
	fh.Write([]byte(xml.Header))

	sh := &sheetEncoder{
		enc: xml.NewEncoder(fh),
	}

	sheetOpen(sh.enc)

	return sh
}

func (sh *sheetEncoder) Close() {
	sheetClose(sh.enc, sh.hyperlinks)
	sh.enc.Flush()
}

func (sh *sheetEncoder) writeRow(cs ...interface{}) {
	var cells []Cell
	for i, v := range cs {
		cell := asValue(v)
		cell.Ref = fmt.Sprintf("%c%d", 'A'+i, sh.rows+1) // FIXME: > 26
		cells = append(cells, cell)

		// hyperlinks refs are written at the end of the sheet
		if link := cell.hyperlink; link != nil {
			linkID := sh.addLinkRelation(link.url)
			sh.hyperlinks = append(sh.hyperlinks, hyperlink{
				RelID:   linkID,
				Ref:     cell.Ref,
				Display: link.Display,
				Tooltip: link.Tooltip,
			})
		}
	}

	sh.enc.EncodeElement(
		rowXML{
			Index: sh.rows + 1,
			Cells: cells,
		},
		xml.StartElement{
			Name: xml.Name{"", "row"},
		},
	)
	sh.rows++
}

func (sh *sheetEncoder) addLinkRelation(url string) string {
	id := fmt.Sprintf("linkId%d", len(sh.relations)+1)
	sh.relations = append(sh.relations, relationship{
		ID:         id,
		Type:       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
		Target:     url,
		TargetMode: "External",
	})
	return id
}

func sheetOpen(enc *xml.Encoder) {
	enc.EncodeToken(xml.StartElement{
		Name: xml.Name{"", "worksheet"},
		Attr: []xml.Attr{
			{xml.Name{"", "xmlns"}, "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
			{xml.Name{"", "xmlns:r"}, "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
		},
	})
	enc.EncodeToken(xml.StartElement{
		Name: xml.Name{"", "sheetData"},
	})
}

func sheetClose(enc *xml.Encoder, links []hyperlink) error {
	if err := enc.EncodeToken(xml.EndElement{
		Name: xml.Name{"", "sheetData"},
	}); err != nil {
		return err
	}

	if err := encodeHyperlinks(enc, links); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.EndElement{
		Name: xml.Name{"", "worksheet"},
	}); err != nil {
		return err
	}
	return nil
}

func encodeHyperlinks(enc *xml.Encoder, links []hyperlink) error {
	if len(links) == 0 {
		return nil
	}
	if err := enc.EncodeToken(xml.StartElement{
		Name: xml.Name{"", "hyperlinks"},
	}); err != nil {
		return err
	}
	for _, link := range links {
		if err := enc.EncodeElement(
			link,
			xml.StartElement{
				Name: xml.Name{"", "hyperlink"},
			},
		); err != nil {
			return err
		}
	}
	return enc.EncodeToken(xml.EndElement{
		Name: xml.Name{"", "hyperlinks"},
	})
}

func asValue(v interface{}) Cell {
	switch vt := v.(type) {
	case int, int64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%d", vt),
		}
	case float64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%f", vt),
		}
	case string:
		return Cell{
			Type:         "inlineStr",
			InlineString: &vt,
		}
	case Cell:
		return vt
	case Hyperlink:
		cell := asValue(vt.Title)
		cell.hyperlink = &hyperlink{
			url:     vt.URL,
			Display: vt.Title,
			Tooltip: vt.Tooltip,
		}
		return cell
	default:
		// FIXME :)
		panic(fmt.Sprintf("unhandled value Fixme! %T", vt))
	}
}

func applyStyle(id int, v interface{}) Cell {
	c := asValue(v)
	c.Style = &id
	return c
}
