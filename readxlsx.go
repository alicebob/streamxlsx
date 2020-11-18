package streamxlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
)

// TestFile is used to test the package.
// It doesn't implement a full reader.
type TestFile struct {
	Sheets []TestSheet
}

type TestSheet struct {
	Name  string
	Cells []TestCell
}

type TestCell struct {
	Ref   string
	Type  string
	Value string
	Style int
}

func TestParse(b []byte) (*TestFile, error) {
	z, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, err
	}

	file := TestFile{}

	var rels relationshipsXML
	if err := readXML(z, "xl/_rels/workbook.xml.rels", &rels); err != nil {
		return nil, err
	}
	var workbook workbookXML
	if err := readXML(z, "xl/workbook.xml", &workbook); err != nil {
		return nil, err
	}

	for i, sheet := range workbook.Sheets {
		s, err := readSheet(z, i+1, sheet.Name)
		if err != nil {
			return nil, err
		}
		file.Sheets = append(file.Sheets, *s)
	}

	return &file, nil
}

func readSheet(z *zip.Reader, id int, name string) (*TestSheet, error) {
	var s worksheetXML
	if err := readXML(z, fmt.Sprintf("xl/worksheets/sheet%d.xml", id), &s); err != nil {
		return nil, err
	}
	var cells []TestCell
	for _, row := range s.Rows {
		for _, cell := range row.Cells {
			v := cell.Value
			if cell.InlineString != nil {
				v = *cell.InlineString
			}
			style := 0
			if cell.Style != nil {
				style = *cell.Style
			}
			cells = append(cells, TestCell{
				Ref:   cell.Ref,
				Type:  cell.Type,
				Value: v,
				Style: style,
			})
		}
	}
	return &TestSheet{
		Name:  name,
		Cells: cells,
	}, nil
}

func readXML(z *zip.Reader, filename string, dest interface{}) error {
	for _, f := range z.File {
		if f.Name == filename {
			fh, err := f.Open()
			if err != nil {
				return err
			}
			return xml.NewDecoder(fh).Decode(dest)
		}
	}
	return errors.New("file not found")
}
