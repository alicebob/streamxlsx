package streamxlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

type workbookXML struct {
	XMLName string  `xml:"workbook"`
	XMLNS   string  `xml:"xmlns,attr"`
	XMLNSR  string  `xml:"xmlns:r,attr"`
	Sheets  []Sheet `xml:"sheets>sheet"`
}

type Sheet struct {
	Name string `xml:"name,attr"`
	ID   string `xml:"sheetId,attr"`
	RID  string `xml:"r:id,attr"`
}

func writeWorkbook(fh io.Writer, sheetTitles []string) error {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	var sheets []Sheet
	for i, title := range sheetTitles {
		sheets = append(sheets, Sheet{
			Name: title,
			ID:   fmt.Sprintf("%d", i+1),
			RID:  fmt.Sprintf("sheetId%d", i+1),
		})
	}

	return enc.Encode(workbookXML{
		XMLNS:  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		XMLNSR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Sheets: sheets,
	})
}
