package streamxlsx

import (
	"bufio"
	"encoding/base64"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"time"
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

func writeCell(w *bufio.Writer, c Cell) error {
	w.WriteString(`<c r="`)
	w.WriteString(c.Ref) // no enc
	if c.Type != "" {
		w.WriteString(`" t="`)
		w.WriteString(c.Type) // no enc
	}
	if c.Style != nil {
		w.WriteString(`" s="`)
		w.WriteString(strconv.Itoa(*c.Style))
	}
	w.WriteString(`">`)
	if c.InlineString != nil {
		w.WriteString(`<is><t>`)
		xml.EscapeText(w, []byte(*c.InlineString))
		w.WriteString(`</t></is>`)
	} else {
		w.WriteString(`<v>`)
		xml.EscapeText(w, []byte(c.Value))
		w.WriteString(`</v>`)
	}
	w.WriteString(`</c>`)
	return nil
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
	buf        *bufio.Writer
	rows       int
	hyperlinks []hyperlink
	relations  []relationship
}

func newSheetEncoder(fh io.Writer) (*sheetEncoder, error) {
	fh.Write([]byte(xml.Header))

	sh := &sheetEncoder{
		buf: bufio.NewWriterSize(fh, 1_000_000),
	}

	return sh, sheetOpen(sh.buf)
}

func (sh *sheetEncoder) Close() {
	sheetClose(sh.buf, sh.hyperlinks)
	sh.buf.Flush()
}

func (sh *sheetEncoder) writeRow(cs ...interface{}) error {
	// var cells []Cell
	fmt.Fprintf(sh.buf, `<row r="%d">`, sh.rows+1)
	for i, v := range cs {
		cell, err := asCell(v)
		if err != nil {
			return err
		}
		cell.Ref = AsRef(i, sh.rows)
		writeCell(sh.buf, cell)

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
	sh.buf.WriteString(`</row>`)

	sh.rows++

	return nil
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

func sheetOpen(w *bufio.Writer) error {
	w.WriteString(`<worksheet
xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
><sheetData>`)
	return nil
}

func sheetClose(w *bufio.Writer, links []hyperlink) error {
	w.WriteString(`</sheetData>`)
	if err := encodeHyperlinks(w, links); err != nil {
		return err
	}
	w.WriteString(`</worksheet>`)
	return nil
}

func encodeHyperlinks(w *bufio.Writer, links []hyperlink) error {
	if len(links) == 0 {
		return nil
	}
	w.WriteString(`<hyperlinks>`)
	for _, link := range links {
		w.WriteString(`<hyperlink r:id="`)
		xml.EscapeText(w, []byte(link.RelID))
		w.WriteString(`" ref="`)
		xml.EscapeText(w, []byte(link.Ref))
		w.WriteString(`" display="`)
		xml.EscapeText(w, []byte(link.Display))
		w.WriteString(`" tooltip="`)
		xml.EscapeText(w, []byte(link.Tooltip))
		w.WriteString(`"/>`)
	}
	w.WriteString(`</hyperlinks>`)
	return nil
}

func asCell(v interface{}) (Cell, error) {
	switch vt := v.(type) {
	case Cell:
		return vt, nil
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%d", vt),
		}, nil
	case float32, float64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%f", vt),
		}, nil
	case []byte:
		return asCell(base64.StdEncoding.EncodeToString(vt))
	case string:
		return Cell{
			Type:         "inlineStr",
			InlineString: &vt,
		}, nil
	case time.Time:
		return Cell{
			Type:  "n",
			Value: oaDate(vt),
		}, nil
	case Hyperlink:
		cell, err := asCell(vt.Title)
		cell.hyperlink = &hyperlink{
			url:     vt.URL,
			Display: vt.Title,
			Tooltip: vt.Tooltip,
		}
		return cell, err
	default:
		return Cell{}, fmt.Errorf("unsupported cell type: %T", vt)
	}
}

func applyStyle(id int, v interface{}) (Cell, error) {
	c, err := asCell(v)
	c.Style = &id
	return c, err
}

// AsRef makes an 'A13' style ref. Arguments are 0-based.
func AsRef(column, row int) string {
	return asCol(column) + strconv.Itoa(row+1)
}

// col number as 'ABC' column ref
func asCol(n int) string {
	// taken from https://github.com/psmithuk/xlsx/blob/master/xlsx.go
	var s string
	n += 1

	for n > 0 {
		n -= 1
		s = string('A'+(n%26)) + s
		n /= 26
	}

	return s
}

func oaDate(d time.Time) string {
	// taken from https://github.com/psmithuk/xlsx/blob/master/xlsx.go
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	nsPerDay := 24 * time.Hour

	v := -1 * float64(epoch.Sub(d)) / float64(nsPerDay)

	// TODO: deal with dates before epoch
	// e.g. http://stackoverflow.com/questions/15549823/oadate-to-milliseconds-timestamp-in-javascript/15550284#15550284

	if d.Hour() == 0 && d.Minute() == 0 && d.Second() == 0 {
		return fmt.Sprintf("%d", int64(v))
	} else {
		return fmt.Sprintf("%f", v)
	}
}
