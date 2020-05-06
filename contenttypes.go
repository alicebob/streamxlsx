package streamxlsx

import (
	"io"
)

func writeContentTypes(fh io.Writer, sheets []string) {
	fh.Write([]byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"></Override>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"></Override>
<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"></Override>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"></Override>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml">
</Override><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
<Default Extension="xml" ContentType="application/xml"></Default>
</Types>`))
}
