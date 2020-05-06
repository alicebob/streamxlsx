package streamxlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

type Row struct {
	Cells []Cell `xml:"c"`
}

type Cell struct {
	Value string `xml:"v"`
	Ref   string `xml:"r,attr"` // "A1" &c.
}

type sheetEncoder struct {
	enc  *xml.Encoder
	rows int
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
	sheetClose(sh.enc)
	sh.enc.Flush()
}

func (sh *sheetEncoder) writeRow(v string) {
	sh.enc.EncodeElement(
		Row{
			Cells: []Cell{
				{
					Value: v,
					Ref:   fmt.Sprintf("%c%d", 'A'+sh.rows, 1),
				},
			},
		},
		xml.StartElement{
			Name: xml.Name{"", "row"},
		},
	)
	sh.rows++
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

func sheetClose(enc *xml.Encoder) {
	if err := enc.EncodeToken(xml.EndElement{
		Name: xml.Name{"", "sheetData"},
	}); err != nil {
		panic(err)
	}
	if err := enc.EncodeToken(xml.EndElement{
		Name: xml.Name{"", "worksheet"},
	}); err != nil {
		panic(err)
	}
}
