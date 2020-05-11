// Package streamxlsx implements a streaming .xlsx (Excel spreadsheat) file writer.
//
// The focus is to easily generate tabular data files.
package streamxlsx

import (
	"archive/zip"
	"bufio"
	"fmt"
	"io"
)

type StreamXLSX struct {
	buf            *bufio.Writer
	zip            *zip.Writer
	openSheet      *sheetEncoder
	finishedSheets []string
	// The stylesheet will be written on Close(). You generally won't want to
	// use this directly, but via `Format()`.
	Styles     *Stylesheet
	styleCache map[string]int
}

// New creates a new file. Do Close() it afterwards.
//
// A StreamXLSX is currently not safe to use with multiple go routines at the same time.
func New(w io.Writer) *StreamXLSX {
	buf := bufio.NewWriterSize(w, 1*1024*1024)
	s := &StreamXLSX{
		buf:        buf,
		zip:        zip.NewWriter(buf),
		Styles:     &Stylesheet{},
		styleCache: map[string]int{},
	}

	// empty style. Not 100% it's needed
	styleID := s.Styles.GetCellStyleID(Xf{})
	s.Styles.GetCellID(Xf{XfID: &styleID})

	s.writeRelations()
	return s
}

// Write a row to the current sheet.
// No values is a valid (empty) row, and not every row needs to have the same
// number of elements.
//
// Supported cell datatypes:
//    all ints and uints, floats, string
// Additional special cases:
//    []byte: will be base64 encoded
//    time.Time: handled, but you need to Format() it. For example: s.Format("mm-dd-yy", aTimeTime)
//    Hyperlink{}: will make the cell a hyperlink
//    Cell{}: if you want to set everything manually
//
// See Format() to apply number formatting to cells.
func (s *StreamXLSX) WriteRow(vs ...interface{}) error {
	sh := s.sheet()
	return sh.writeRow(vs...)
}

// WriteSheet closes the currenly open sheet, with the given title.
// The process is you first do all the `WriteRow()`s for a sheet, followed by
// its WriteSheet().  There is always an open sheet. You don't have to close
// the final sheet, but it'll give you a boring name ("sheet N").
func (s *StreamXLSX) WriteSheet(title string) error {
	s.openSheet.Close()
	if err := s.writeSheetRelations(); err != nil { // for hyperlink refs
		return err
	}
	s.openSheet = nil
	s.finishedSheets = append(s.finishedSheets, title)
	return nil
}

// Adds a number format to a cell. Examples or formats are "0.00", "0%", ...
// This is used to wrap a value in a WriteRow().
func (s *StreamXLSX) Format(code string, cell interface{}) Cell {
	if xfID, ok := s.styleCache[code]; ok {
		c, _ := applyStyle(xfID, cell) // FIXME: error is ignored
		return c
	}

	numFmtID := s.Styles.GetNumFmtID(code)
	cellStyleID := s.Styles.GetCellStyleID(Xf{})
	styleFx := Xf{
		NumFmtID:          numFmtID,
		ApplyNumberFormat: 1,
		XfID:              &cellStyleID,
	}
	xfID := s.Styles.GetCellID(styleFx)
	s.styleCache[code] = xfID
	c, _ := applyStyle(xfID, cell) // FIXME: error is ignored
	return c
}

// Adds a hyperlink in a cell. You can use these as a value in WriteRow().
// (implementation detail: parts of the hyperlink datastructure is
// only written when closing a sheet, so they are buffered)
type Hyperlink struct {
	URL, Title, Tooltip string
}

// Finish writing the spreadsheet.
func (s *StreamXLSX) Close() error {
	if s.openSheet != nil {
		if err := s.WriteSheet(fmt.Sprintf("sheet %d", len(s.finishedSheets)+1)); err != nil {
			return err
		}

	}

	if err := s.writeWorkbook(); err != nil {
		return err
	}
	if err := s.writeStylesheet(); err != nil {
		return err
	}
	if err := s.writeWorkbookRelations(); err != nil {
		return err
	}
	if err := s.writeContentTypes(); err != nil {
		return err
	}
	if err := s.zip.Close(); err != nil {
		return err
	}
	return s.buf.Flush()
}

func (s *StreamXLSX) sheet() *sheetEncoder {
	if s.openSheet != nil {
		return s.openSheet
	}
	filename := fmt.Sprintf("xl/worksheets/sheet%d.xml", len(s.finishedSheets)+1)
	fh, _ := s.zip.Create(filename) // no need to close!

	enc, _ := newSheetEncoder(fh) // FIXME err

	s.openSheet = enc
	return s.openSheet
}

func (s *StreamXLSX) writeWorkbook() error {
	filename := "xl/workbook.xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeWorkbook(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeStylesheet() error {
	filename := "xl/styles.xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeStylesheet(fh, s.Styles)
}

func (s *StreamXLSX) writeRelations() error {
	filename := "_rels/.rels"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeRelations(fh)
}

func (s *StreamXLSX) writeWorkbookRelations() error {
	filename := "xl/_rels/workbook.xml.rels"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeWorkbookRelations(fh, s.finishedSheets)
}

func (s *StreamXLSX) writeSheetRelations() error {
	if len(s.openSheet.relations) == 0 {
		return nil
	}
	filename := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", len(s.finishedSheets)+1)
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeRelations_(fh, s.openSheet.relations)
}

func (s *StreamXLSX) writeContentTypes() error {
	filename := "[Content_Types].xml"
	fh, err := s.zip.Create(filename)
	if err != nil {
		return err
	}
	return writeContentTypes(fh, len(s.finishedSheets))
}
