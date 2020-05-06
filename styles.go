package streamxlsx

import (
	"encoding/xml"
	"io"
)

type Stylesheet struct {
	NumFmts      []NumFmt
	CellXfs      []Xf
	CellStyleXfs []Xf
}

type StylesheetXML struct {
	XMLName      string     `xml:"styleSheet"`
	XMLNS        string     `xml:"xmlns,attr"`
	NumFmts      NumFmtsXML `xml:"numFmts"`
	Fonts        FontsXML   `xml:"fonts"`
	Fills        FillsXML   `xml:"fills"`
	Borders      BordersXML `xml:"borders"`
	CellStyleXfs XfsXML     `xml:"cellStyleXfs"`
	CellXfs      XfsXML     `xml:"cellXfs"`
}

type NumFmtsXML struct {
	Count   int      `xml:"count,attr"`
	NumFmts []NumFmt `xml:"numFmt"`
}

type FontsXML struct {
	Count int    `xml:"count,attr"`
	Fonts []Font `xml:"font"`
}

/*
type FontSZ struct {
	Val int `xml:"val,attr"`
}
type FontColor struct {
	Theme int `xml:"theme,attr"`
}
type FontName struct {
	Val string `xml:"val,attr"`
}
type FontFamily struct {
	Val int `xml:"val,attr"`
}
type FontScheme struct {
	Val string `xml:"val,attr"`
}
*/
type Font struct {
	// SZ     FontSZ     `xml:"sz"`
	// Color  FontColor  `xml:"color"`
	// Name   FontName   `xml:"name"`
	// Family FontFamily `xml:"family"`
	// Scheme FontScheme `xml:"scheme"`
}

type FillPattern struct {
	Type string `xml:"patternType,attr"`
}
type Fill struct {
	PatternFill FillPattern `xml:"patternFill"`
}
type FillsXML struct {
	Count int    `xml:"count,attr"`
	Fills []Fill `xml:"fill"`
}

type BordersXML struct {
	Count   int      `xml:"count,attr"`
	Borders []Border `xml:"border"`
}

type Border struct {
	// <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
}

type XfsXML struct {
	Count int  `xml:"count,attr"`
	Xfs   []Xf `xml:"xf"`
}

// <xf numFmtId="0" fontId="8" fillId="4" borderId="0" xfId="3"/>
type Xf struct {
	NumFmtID          int  `xml:"numFmtId,attr"`
	FontID            int  `xml:"fontId,attr"`
	FillID            int  `xml:"fillId,attr"`
	BorderID          int  `xml:"borderId,attr"`
	ApplyNumberFormat int  `xml:"applyNumberFormat,attr,omitempty"`
	XfID              *int `xml:"xfId,attr,omitempty"`
}

type NumFmt struct {
	ID   int    `xml:"numFmtId,attr"`
	Code string `xml:"formatCode,attr"`
}

// get or create the ID for a numfmt. It can return a "default" ID, or create a custom ID.
func (s *Stylesheet) getNumFmtID(code string) int {
	// see builtInNumFmt in tealeg
	switch code {
	case "0":
		return 1
	case "0.00":
		return 2
	case "#.##0":
		return 3
	}
	// FIXME &c.

	max := 163 // custom IDs start here, according to tealeg
	for _, nf := range s.NumFmts {
		if nf.Code == code {
			return nf.ID
		}
		if nf.ID > max {
			max = nf.ID
		}
	}
	newID := max + 1
	s.NumFmts = append(s.NumFmts, NumFmt{
		ID:   newID,
		Code: code,
	})
	return newID
}

// makes a CellXF ID
// The ID is the entry in the array, 0-based
func (s *Stylesheet) getCellID(xf Xf) int {
	for i, x := range s.CellXfs {
		if x == xf {
			return i
		}
	}
	s.CellXfs = append(s.CellXfs, xf)
	return len(s.CellXfs) - 1
}

// makes a CellStyleXf ID
// The ID is the entry in the array, 0-based
func (s *Stylesheet) getCellStyleID(xf Xf) int {
	for i, x := range s.CellStyleXfs {
		if x == xf {
			return i
		}
	}
	s.CellStyleXfs = append(s.CellStyleXfs, xf)
	return len(s.CellStyleXfs) - 1
}

func writeStylesheet(fh io.Writer, s *Stylesheet) {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	enc.Encode(StylesheetXML{
		XMLNS: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		NumFmts: NumFmtsXML{
			Count:   len(s.NumFmts),
			NumFmts: s.NumFmts,
		},
		Fonts: FontsXML{
			Count: 1,
			Fonts: []Font{
				{
					// SZ:     FontSZ{11},
					// Color:  FontColor{1},
					// Name:   FontName{"Calibri"},
					// Family: FontFamily{2},
					// Scheme: FontScheme{"minor"},
				},
			},
		},
		Fills: FillsXML{
			Count: 1,
			Fills: []Fill{
				{PatternFill: FillPattern{Type: "none"}},
			},
		},
		Borders: BordersXML{
			Count: 1,
			Borders: []Border{
				{},
			},
		},
		CellStyleXfs: XfsXML{
			Count: len(s.CellStyleXfs),
			Xfs:   s.CellStyleXfs,
		},
		CellXfs: XfsXML{
			Count: len(s.CellXfs),
			Xfs:   s.CellXfs,
		},
	})
}
