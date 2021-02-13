package streamxlsx

import (
	"io"
)

// we don't use shared strings, but we make an empty file just in case some
// excel version expects one
func writeSharedStrings(fh io.Writer) error {
	_, err := fh.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>
`))
	return err
}
