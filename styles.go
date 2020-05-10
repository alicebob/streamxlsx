package streamxlsx

import (
	"encoding/xml"
	"io"
)

// A Stylesheet has all used formats and styles. There is exactly one per document.
// It's recommended to use `Format()` to work with styles, since that hides all the details.
//
// note: this could have support for fonts, fills, and borders.
type Stylesheet struct {
	NumFmts      []NumFmt
	CellXfs      []Xf
	CellStyleXfs []Xf
}

type stylesheetXML struct {
	XMLName      string     `xml:"styleSheet"`
	XMLNS        string     `xml:"xmlns,attr"`
	NumFmts      numFmtsXML `xml:"numFmts"`
	Fonts        fontsXML   `xml:"fonts"`
	Fills        fillsXML   `xml:"fills"`
	Borders      bordersXML `xml:"borders"`
	CellStyleXfs xfsXML     `xml:"cellStyleXfs"`
	CellXfs      xfsXML     `xml:"cellXfs"`
}

type numFmtsXML struct {
	Count   int      `xml:"count,attr"`
	NumFmts []NumFmt `xml:"numFmt"`
}

type fontsXML struct {
	Count int    `xml:"count,attr"`
	Fonts []font `xml:"font"`
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
// not implemented
type font struct {
	// SZ     FontSZ     `xml:"sz"`
	// Color  FontColor  `xml:"color"`
	// Name   FontName   `xml:"name"`
	// Family FontFamily `xml:"family"`
	// Scheme FontScheme `xml:"scheme"`
}

type fillPattern struct {
	Type string `xml:"patternType,attr"`
}

// not implemented
type fill struct {
	PatternFill fillPattern `xml:"patternFill"`
}
type fillsXML struct {
	Count int    `xml:"count,attr"`
	Fills []fill `xml:"fill"`
}

type bordersXML struct {
	Count   int      `xml:"count,attr"`
	Borders []border `xml:"border"`
}

// not implemented
type border struct {
	// <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
}

type xfsXML struct {
	Count int  `xml:"count,attr"`
	Xfs   []Xf `xml:"xf"`
}

// Xf is either a CellXF or a CellStyleXf.
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

// Get or create the ID for a numfmt. It can return a "default" ID, or create a
// custom ID.
// Example of a code is "0.00".
func (s *Stylesheet) GetNumFmtID(code string) int {
	// see builtInNumFmt in tealeg
	switch code {
	case "0":
		return 1
	case "0.00":
		return 2
	case "#.##0":
		return 3
	default:
		// FIXME &c.
	}

	max := 163 // custom IDs start here+1, according to tealeg
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
func (s *Stylesheet) GetCellID(xf Xf) int {
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
func (s *Stylesheet) GetCellStyleID(xf Xf) int {
	for i, x := range s.CellStyleXfs {
		if x == xf {
			return i
		}
	}
	s.CellStyleXfs = append(s.CellStyleXfs, xf)
	return len(s.CellStyleXfs) - 1
}

func writeStylesheet(fh io.Writer, s *Stylesheet) error {
	fh.Write([]byte(xml.Header))
	enc := xml.NewEncoder(fh)

	return enc.Encode(stylesheetXML{
		XMLNS: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		NumFmts: numFmtsXML{
			Count:   len(s.NumFmts),
			NumFmts: s.NumFmts,
		},
		Fonts: fontsXML{
			Count: 1,
			Fonts: []font{
				{
					// SZ:     FontSZ{11},
					// Color:  FontColor{1},
					// Name:   FontName{"Calibri"},
					// Family: FontFamily{2},
					// Scheme: FontScheme{"minor"},
				},
			},
		},
		Fills: fillsXML{
			Count: 1,
			Fills: []fill{
				{PatternFill: fillPattern{Type: "none"}},
			},
		},
		Borders: bordersXML{
			Count: 1,
			Borders: []border{
				{},
			},
		},
		CellStyleXfs: xfsXML{
			Count: len(s.CellStyleXfs),
			Xfs:   s.CellStyleXfs,
		},
		CellXfs: xfsXML{
			Count: len(s.CellXfs),
			Xfs:   s.CellXfs,
		},
	})
}
