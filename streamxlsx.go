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
	styles         *Stylesheet
}

func New(w io.Writer) *StreamXLSX {
	s := &StreamXLSX{
		w:      w,
		zip:    zip.NewWriter(w),
		styles: &Stylesheet{},
	}
	// empty style. Not 100% it's needed
	styleID := s.styles.getCellStyleID(Xf{})
	s.styles.getCellID(Xf{XfID: &styleID})

	s.writeRelations()
	return s
}

// not all type are supported yet.
// Don't write more than 26 columns :)
func (s *StreamXLSX) WriteRow(vs ...interface{}) {
	sh := s.sheet()
	sh.writeRow(vs...)
}

func (s *StreamXLSX) WriteSheet(title string) {
	s.openSheet.Close()
	s.writeSheetRelations() // for hyperlink refs
	s.openSheet = nil
	s.finishedSheets = append(s.finishedSheets, title)
}

// Adds a number format to a cell. Examples are "0.00", "0%", ...
func (s *StreamXLSX) Format(code string, cell interface{}) Cell {
	numFmtID := s.styles.getNumFmtID(code)
	cellStyleID := s.styles.getCellStyleID(Xf{})
	styleFx := Xf{NumFmtID: numFmtID, ApplyNumberFormat: 1, XfID: &cellStyleID}
	xfID := s.styles.getCellID(styleFx)
	return Styled(xfID, cell)
}

// Adds a hyperlink ref. They are streamed only at the end of every sheet.
type Hyperlink struct {
	URL, Title, Tooltip string
}

func (s *StreamXLSX) Close() {
	if s.openSheet != nil {
		s.WriteSheet(fmt.Sprintf("sheet %d", len(s.finishedSheets)+1))
	}

	s.writeWorkbook()
	s.writeStylesheet()
	s.writeWorkbookRelations()
	s.writeContentTypes()
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

func (s *StreamXLSX) writeWorkbook() {
	filename := "xl/workbook.xml"
	fh, _ := s.zip.Create(filename)
	writeWorkbook(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeStylesheet() {
	filename := "xl/styles.xml"
	fh, _ := s.zip.Create(filename)
	writeStylesheet(fh, s.styles)
}

func (s *StreamXLSX) writeRelations() {
	filename := "_rels/.rels"
	fh, _ := s.zip.Create(filename)
	writeRelations(fh)
}

func (s *StreamXLSX) writeWorkbookRelations() {
	filename := "xl/_rels/workbook.xml.rels"
	fh, _ := s.zip.Create(filename)
	writeWorkbookRelations(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeSheetRelations() {
	if len(s.openSheet.relations) == 0 {
		return
	}
	filename := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", len(s.finishedSheets)+1)
	fh, _ := s.zip.Create(filename)
	writeRelations_(fh, s.openSheet.relations)
}

func (s *StreamXLSX) writeContentTypes() {
	filename := "[Content_Types].xml"
	fh, _ := s.zip.Create(filename)
	writeContentTypes(fh, len(s.finishedSheets))
}
