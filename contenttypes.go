package streamxlsx

import (
	"fmt"
	"io"
)

func writeContentTypes(fh io.Writer, sheetCount int) {
	fh.Write([]byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override>
`))
	for i := 0; i < sheetCount; i++ {
		fmt.Fprintf(fh, `<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>`,
			i+1,
		)
		fmt.Fprintf(fh, `<Override PartName="/xl/worksheets/_rels/sheet%d.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>`,
			i+1,
		)
	}
	fh.Write([]byte(`</Types>`))
}
