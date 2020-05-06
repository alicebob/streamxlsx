package streamxlsx_test

import (
	"bytes"
	"testing"

	"github.com/alicebob/streamxlsx"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/tealeg/xlsx/v2"
)

func TestBasic(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("12", "<-- that's a string", "next is an int -->", 22)
	s.WriteRow("aap")
	s.WriteRow("noot")
	s.WriteRow("mies")
	s.WriteRow("vuur")
	s.WriteSheet("that was sheet 1")

	s.WriteRow("13")
	s.WriteSheet("that was sheet 2")
	s.Close()

	// Read it back again
	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 2)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "that was sheet 1", sheet0.Name)
	require.Len(t, sheet0.Rows, 5)
	require.Len(t, sheet0.Rows[0].Cells, 4)

	cell00 := sheet0.Rows[0].Cells[0]
	cell03 := sheet0.Rows[0].Cells[3]
	assert.Equal(t, "12", cell00.String())
	assert.Equal(t, xlsx.CellTypeInline, cell00.Type())
	assert.Equal(t, "22", cell03.String())
	assert.Equal(t, xlsx.CellTypeNumeric, cell03.Type())

	sheet1 := xf.Sheets[1]
	require.Equal(t, "that was sheet 2", sheet1.Name)
	require.Len(t, sheet1.Rows, 1)
	require.Len(t, sheet1.Rows[0].Cells, 1)
	assert.Equal(t, "13", sheet1.Rows[0].Cells[0].String())

}

func TestDatatypes(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("a string", "hello world!")
	s.WriteRow("also a string", "14")
	s.WriteRow("a number", 14)
	s.WriteRow("a float", 3.1415)
	s.WriteRow("a styled float (default)", s.Format("0.00", 3.1415))
	s.WriteRow("a styled float (custom)", s.Format("0.000", 3.1415))
	// s.WriteRow("a link", Hyperlink("http://example.com", "clickme", "I'm a tooltip"))
	s.Close()

	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Len(t, sheet0.Rows, 6)

	test := func(cell *xlsx.Cell, typ xlsx.CellType, value string) {
		t.Helper()
		assert.Equal(t, typ, cell.Type())
		assert.Equal(t, value, cell.String())
	}
	test(sheet0.Rows[0].Cells[1], xlsx.CellTypeInline, "hello world!")
	test(sheet0.Rows[1].Cells[1], xlsx.CellTypeInline, "14")
	test(sheet0.Rows[2].Cells[1], xlsx.CellTypeNumeric, "14")
	test(sheet0.Rows[3].Cells[1], xlsx.CellTypeNumeric, "3.1415")
	test(sheet0.Rows[4].Cells[1], xlsx.CellTypeNumeric, "3.14")
	test(sheet0.Rows[5].Cells[1], xlsx.CellTypeNumeric, "3.142")
	// test(sheet0.Rows[4].Cells[1], xlsx.CellTypeInline, "http://example.com")
}

func TestHangingSheet(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("aap")
	s.WriteRow("noot")
	// no s.WriteSheet
	s.Close()

	// Read it back again
	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Rows, 2)
}
