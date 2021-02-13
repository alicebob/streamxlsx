package streamxlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

type relationshipsXML struct {
	XMLName       string         `xml:"Relationships"`
	XMLNS         string         `xml:"xmlns,attr"`
	Relationships []relationship `xml:"Relationship"`
}

type relationship struct {
	ID         string `xml:"Id,attr"`
	Target     string `xml:"Target,attr"`
	Type       string `xml:"Type,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

func writeRelations(fh io.Writer) error {
	return writeRelations_(fh, []relationship{
		{
			ID:     "rId1",
			Target: "/xl/workbook.xml",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		},
	})
}

func writeWorkbookRelations(fh io.Writer, sheetTitles []string) error {
	var rels = []relationship{
		{
			ID:     "style1",
			Target: "/xl/styles.xml",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		},
		{
			ID:     "shared1",
			Target: "/xl/sharedStrings.xml",
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
		},
	}
	for i := range sheetTitles {
		rels = append(rels, relationship{
			ID:     fmt.Sprintf("sheetId%d", i+1),
			Target: fmt.Sprintf("/xl/worksheets/sheet%d.xml", i+1),
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		})
	}
	return writeRelations_(fh, rels)
}

func writeRelations_(fh io.Writer, rels []relationship) error {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	return enc.Encode(relationshipsXML{
		XMLNS:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: rels,
	})
}
