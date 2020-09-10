package streamxlsx

import (
	"bufio"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
)

type rowXML struct {
	Cells []Cell `xml:"c"`
	Index int    `xml:"r,attr"`
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
	fmt.Fprintf(sh.buf, `<row r="%d">`, sh.rows+1)
	for i, v := range cs {
		if v == nil {
			continue
		}
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
		s = string(rune('A'+(n%26))) + s
		n /= 26
	}

	return s
}
