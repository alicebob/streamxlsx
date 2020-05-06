package streamxlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

type Relationships struct {
	XMLName       string         `xml:"Relationships"`
	XMLNS         string         `xml:"xmlns,attr"`
	Relationships []Relationship `xml:"Relationship"`
}

type Relationship struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
	Type   string `xml:"Type,attr"`
}

func writeRelations(fh io.Writer) {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	enc.Encode(Relationships{
		XMLNS: "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: []Relationship{
			{
				ID:     "rId1",
				Target: "xl/workbook.xml",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
			},
		},
	})
}

func writeXLRelations(fh io.Writer, sheetTitles []string) {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	var rels []Relationship
	for i := range sheetTitles {
		rels = append(rels, Relationship{
			ID:     fmt.Sprintf("sheetId%d", i+1),
			Target: fmt.Sprintf("worksheets/sheet%d.xml", i+1), // FIXME
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		})
	}

	enc.Encode(Relationships{
		XMLNS:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: rels,
	})
}
