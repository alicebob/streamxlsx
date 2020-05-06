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
	ID         string `xml:"Id,attr"`
	Target     string `xml:"Target,attr"`
	Type       string `xml:"Type,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

func writeRelations(fh io.Writer) {
	writeRelations_(fh, []Relationship{
		{
			ID:     "rId1",
			Target: "/xl/workbook.xml",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		},
	})
}

func writeWorkbookRelations(fh io.Writer, sheetTitles []string) error {
	var rels = []Relationship{
		{
			ID:     "style1",
			Target: "/xl/styles.xml",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		},
	}
	for i := range sheetTitles {
		rels = append(rels, Relationship{
			ID:     fmt.Sprintf("sheetId%d", i+1),
			Target: fmt.Sprintf("/xl/worksheets/sheet%d.xml", i+1),
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		})
	}
	return writeRelations_(fh, rels)
}

func writeRelations_(fh io.Writer, rels []Relationship) error {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	return enc.Encode(Relationships{
		XMLNS:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: rels,
	})
}
