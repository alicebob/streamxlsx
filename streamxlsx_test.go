package streamxlsx_test

import (
	"bytes"
	"os"
	"testing"
	"time"

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

	// strings
	require.NoError(t, s.WriteRow("a string", "hello world!"))
	require.NoError(t, s.WriteRow("a char", 'q'))
	require.NoError(t, s.WriteRow("bytes", []byte("hi there")))
	require.NoError(t, s.WriteSheet("strings"))

	// numbers
	require.NoError(t, s.WriteRow("a number", 14))
	require.NoError(t, s.WriteRow("a number", int(15)))
	require.NoError(t, s.WriteRow("a number", int8(16)))
	require.NoError(t, s.WriteRow("a number", int16(17)))
	require.NoError(t, s.WriteRow("a number", int32(18)))
	require.NoError(t, s.WriteRow("a number", int64(19)))
	require.NoError(t, s.WriteRow("a number", uint(995)))
	require.NoError(t, s.WriteRow("a number", uint8(99)))
	require.NoError(t, s.WriteRow("a number", uint16(997)))
	require.NoError(t, s.WriteRow("a number", uint32(998)))
	require.NoError(t, s.WriteRow("a number", uint64(999)))
	require.NoError(t, s.WriteRow("a float", 3.1415))
	require.NoError(t, s.WriteRow("a float", float32(3.1415)))
	require.NoError(t, s.WriteRow("a float", float64(3.1415)))
	require.NoError(t, s.WriteSheet("numbers"))

	// misc
	require.NoError(t, s.WriteRow("a link", streamxlsx.Hyperlink{"http://example.com", "clickme", "I'm a tooltip"}))
	require.NoError(t, s.WriteRow("a datetime", s.Format(streamxlsx.DefaultDatetimeFormat, time.Date(2010, 10, 10, 10, 10, 10, 0, time.UTC))))
	require.NoError(t, s.WriteRow("bool", true, false))
	require.NoError(t, s.WriteSheet("misc"))

	s.Close()

	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 3)

	t.Run("string values", func(t *testing.T) {
		test := func(cell *xlsx.Cell, typ xlsx.CellType, value string) {
			t.Helper()
			assert.Equal(t, typ, cell.Type())
			assert.Equal(t, value, cell.String())
		}

		strings := xf.Sheets[0]
		require.Len(t, strings.Rows, 3)
		test(strings.Rows[0].Cells[1], xlsx.CellTypeInline, "hello world!")
		test(strings.Rows[1].Cells[1], xlsx.CellTypeNumeric, "113")
		test(strings.Rows[2].Cells[1], xlsx.CellTypeInline, "aGkgdGhlcmU=")
	})

	t.Run("numeric values", func(t *testing.T) {
		test := func(cell *xlsx.Cell, typ xlsx.CellType, value string) {
			t.Helper()
			assert.Equal(t, typ, cell.Type())
			assert.Equal(t, value, cell.String())
		}

		nums := xf.Sheets[1]
		require.Len(t, nums.Rows, 14)
		test(nums.Rows[0].Cells[1], xlsx.CellTypeNumeric, "14")
		test(nums.Rows[1].Cells[1], xlsx.CellTypeNumeric, "15")
		test(nums.Rows[2].Cells[1], xlsx.CellTypeNumeric, "16")
		test(nums.Rows[3].Cells[1], xlsx.CellTypeNumeric, "17")
		test(nums.Rows[4].Cells[1], xlsx.CellTypeNumeric, "18")
		test(nums.Rows[5].Cells[1], xlsx.CellTypeNumeric, "19")
		test(nums.Rows[6].Cells[1], xlsx.CellTypeNumeric, "995")
		test(nums.Rows[7].Cells[1], xlsx.CellTypeNumeric, "99")
		test(nums.Rows[8].Cells[1], xlsx.CellTypeNumeric, "997")
		test(nums.Rows[9].Cells[1], xlsx.CellTypeNumeric, "998")
		test(nums.Rows[10].Cells[1], xlsx.CellTypeNumeric, "999")
		test(nums.Rows[11].Cells[1], xlsx.CellTypeNumeric, "3.1415")
		test(nums.Rows[12].Cells[1], xlsx.CellTypeNumeric, "3.1415")
		test(nums.Rows[13].Cells[1], xlsx.CellTypeNumeric, "3.1415")
	})

	t.Run("numeric values", func(t *testing.T) {
		test := func(cell *xlsx.Cell, typ xlsx.CellType, value string) {
			t.Helper()
			assert.Equal(t, typ, cell.Type())
			assert.Equal(t, value, cell.String())
		}
		miscs := xf.Sheets[2]
		require.Len(t, miscs.Rows, 3)
		test(miscs.Rows[0].Cells[1], xlsx.CellTypeInline, "clickme")
		test(miscs.Rows[1].Cells[1], xlsx.CellTypeNumeric, "10/10/10 10:10")
		test(miscs.Rows[2].Cells[1], xlsx.CellTypeBool, "TRUE")
		test(miscs.Rows[2].Cells[2], xlsx.CellTypeBool, "FALSE")
	})
}

func TestFormats(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("a float", 3.1415)
	s.WriteRow("a styled float (default)", s.Format("0.00", 3.1415))
	s.WriteRow("a styled float (custom)", s.Format("0.000", 3.1415))
	s.Close()

	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Len(t, sheet0.Rows, 3)

	test := func(cell *xlsx.Cell, typ xlsx.CellType, value string) {
		t.Helper()
		assert.Equal(t, typ, cell.Type())
		assert.Equal(t, value, cell.String())
	}
	test(sheet0.Rows[0].Cells[1], xlsx.CellTypeNumeric, "3.1415")
	test(sheet0.Rows[1].Cells[1], xlsx.CellTypeNumeric, "3.14")
	test(sheet0.Rows[2].Cells[1], xlsx.CellTypeNumeric, "3.142")
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

func TestEmptyFile(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.Close()

	// Read it back again
	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Rows, 0)
}

func TestEmptySheet(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteSheet("sheet 1")
	s.WriteSheet("sheet 2")
	s.Close()

	// Read it back again
	xf, err := xlsx.OpenBinary(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 2)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Rows, 0)
}

func TestWriteError(t *testing.T) {
	fh, err := os.Create("/tmp/streamxlsx.test")
	require.NoError(t, err)
	defer fh.Close()
	defer os.Remove("/tmp/streamxlsx.test")

	s := streamxlsx.New(fh)
	s.WriteRow("aap")
	s.WriteRow("noot")
	fh.Close() // !
	s.WriteSheet("sheet 1")
	s.WriteRow("aap")
	s.WriteRow("noot")
	s.WriteSheet("sheet 2")
	s.WriteRow("aap")
	require.EqualError(t, s.Close(), "write /tmp/streamxlsx.test: file already closed")
}
