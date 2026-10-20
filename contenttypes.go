package streamxlsx

import (
	"fmt"
	"io"
)

// imageTypeToContentType returns the MIME type for an image format
func imageTypeToContentType(format string) string {
	switch format {
	case "png":
		return "image/png"
	case "jpg", "jpeg":
		return "image/jpeg"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	default:
		return "application/octet-stream"
	}
}

func writeContentTypes(fh io.Writer, sheetCount int, imageFormats []string, uniqueImageCount int) error {
	fh.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
`))
	for i := 0; i < sheetCount; i++ {
		fmt.Fprintf(fh, `<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>`,
			i+1,
		)
		fmt.Fprintf(fh, `<Override PartName="/xl/worksheets/_rels/sheet%d.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>`,
			i+1,
		)
	}

	// Add image content types
	for _, format := range imageFormats {
		contentType := imageTypeToContentType(format)
		// Add Default for the extension - this covers all files with this extension
		fmt.Fprintf(fh, `<Default Extension="%s" ContentType="%s"/>`, format, contentType)
		// Add explicit Override for each media file (like Excel/LibreOffice does)
		for j := 1; j <= uniqueImageCount; j++ {
			fmt.Fprintf(fh, `<Override PartName="/xl/media/image%d.%s" ContentType="%s"/>`+"\n", j, format, contentType)
		}
	}

	// Add drawing content types for each sheet that might have a drawing
	// We add overrides for all possible drawing files (1 to sheetCount)
	// Only sheets with images will actually have drawing files, but this is safe
	for i := 1; i <= sheetCount; i++ {
		fmt.Fprintf(fh, `<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`+"\n", i)
		fmt.Fprintf(fh, `<Override PartName="/xl/drawings/_rels/drawing%d.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`+"\n", i)
	}

	_, err := fh.Write([]byte(`</Types>`))
	return err
}
