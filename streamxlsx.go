package streamxlsx

import (
	"archive/zip"
	"fmt"
	"io"
)

type StreamXLSX struct {
	w              io.Writer
	zip            *zip.Writer
	openSheet      *sheetEncoder
	finishedSheets []string
}

func New(w io.Writer) *StreamXLSX {
	s := &StreamXLSX{
		w:   w,
		zip: zip.NewWriter(w),
	}
	s.writeRelations()
	return s
}

func (s *StreamXLSX) WriteRow(v string) {
	sh := s.sheet()
	sh.writeRow(v)
}

func (s *StreamXLSX) WriteSheet(title string) {
	s.openSheet.Close()
	s.openSheet = nil // haha
	s.finishedSheets = append(s.finishedSheets, title)
}

func (s *StreamXLSX) Close() {
	// todo: finish open sheet
	s.writeWorkbook()
	s.writeContentTypes()
	s.writeXLRelations()
	s.zip.Close()
}

func (s *StreamXLSX) sheet() *sheetEncoder {
	if s.openSheet != nil {
		return s.openSheet
	}
	filename := fmt.Sprintf("xl/worksheets/sheet%d.xml", len(s.finishedSheets)+1)
	fh, _ := s.zip.Create(filename) // no need to close!

	enc := newSheetEncoder(fh)

	s.openSheet = enc
	return s.openSheet
}

func (s *StreamXLSX) writeContentTypes() {
	filename := "[Content_Types].xml"
	fh, _ := s.zip.Create(filename)
	writeContentTypes(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeWorkbook() {
	filename := "xl/workbook.xml"
	fh, _ := s.zip.Create(filename)
	writeWorkbook(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeRelations() {
	filename := "_rels/.rels"
	fh, _ := s.zip.Create(filename)
	writeRelations(fh)
}

func (s *StreamXLSX) writeXLRelations() {
	filename := "xl/_rels/workbook.xml.rels"
	fh, _ := s.zip.Create(filename)
	writeXLRelations(fh, s.finishedSheets)
}
