package streamxlsx_test

import (
	"bytes"
	"os"
	"testing"
	"time"

	"github.com/alicebob/streamxlsx"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
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
	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 2)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "that was sheet 1", sheet0.Name)
	assert.Equal(t,
		[]streamxlsx.TestCell{
			{"A1", "inlineStr", "12", 0},
			{"B1", "inlineStr", "<-- that's a string", 0},
			{"C1", "inlineStr", "next is an int -->", 0},
			{"D1", "n", "22", 0},
			{"A2", "inlineStr", "aap", 0},
			{"A3", "inlineStr", "noot", 0},
			{"A4", "inlineStr", "mies", 0},
			{"A5", "inlineStr", "vuur", 0},
		},
		sheet0.Cells,
	)
	sheet1 := xf.Sheets[1]
	require.Equal(t, "that was sheet 2", sheet1.Name)
	assert.Equal(t,
		[]streamxlsx.TestCell{
			{"A1", "inlineStr", "13", 0},
		},
		sheet1.Cells,
	)
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

	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 3)

	t.Run("string values", func(t *testing.T) {
		sheet := xf.Sheets[0]
		assert.Equal(t,
			[]streamxlsx.TestCell{
				{"A1", "inlineStr", "a string", 0},
				{"B1", "inlineStr", "hello world!", 0},
				{"A2", "inlineStr", "a char", 0},
				{"B2", "n", "113", 0},
				{"A3", "inlineStr", "bytes", 0},
				{"B3", "inlineStr", "aGkgdGhlcmU=", 0},
			},
			sheet.Cells,
		)
	})

	t.Run("numeric values", func(t *testing.T) {
		sheet := xf.Sheets[1]
		assert.Equal(t,
			[]streamxlsx.TestCell{
				{"A1", "inlineStr", "a number", 0},
				{"B1", "n", "14", 0},
				{"A2", "inlineStr", "a number", 0},
				{"B2", "n", "15", 0},
				{"A3", "inlineStr", "a number", 0},
				{"B3", "n", "16", 0},
				{"A4", "inlineStr", "a number", 0},
				{"B4", "n", "17", 0},
				{"A5", "inlineStr", "a number", 0},
				{"B5", "n", "18", 0},
				{"A6", "inlineStr", "a number", 0},
				{"B6", "n", "19", 0},
				{"A7", "inlineStr", "a number", 0},
				{"B7", "n", "995", 0},
				{"A8", "inlineStr", "a number", 0},
				{"B8", "n", "99", 0},
				{"A9", "inlineStr", "a number", 0},
				{"B9", "n", "997", 0},
				{"A10", "inlineStr", "a number", 0},
				{"B10", "n", "998", 0},
				{"A11", "inlineStr", "a number", 0},
				{"B11", "n", "999", 0},
				{"A12", "inlineStr", "a float", 0},
				{"B12", "n", "3.141500", 0},
				{"A13", "inlineStr", "a float", 0},
				{"B13", "n", "3.141500", 0},
				{"A14", "inlineStr", "a float", 0},
				{"B14", "n", "3.141500", 0},
			},
			sheet.Cells,
		)
	})
	t.Run("numeric values", func(t *testing.T) {
		sheet := xf.Sheets[2]
		assert.Equal(t,
			[]streamxlsx.TestCell{
				{"A1", "inlineStr", "a link", 0},
				{"B1", "inlineStr", "clickme", 0},
				{"A2", "inlineStr", "a datetime", 0},
				{"B2", "n", "40461.423727", 1},
				{"A3", "inlineStr", "bool", 0},
				{"B3", "b", "1", 0},
				{"C3", "b", "0", 0},
			},
			sheet.Cells,
		)
	})
}

func TestFormats(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("a float", 3.1415)
	s.WriteRow("a styled float (default)", s.Format("0.00", 3.1415))
	s.WriteRow("a styled float (custom)", s.Format("0.000", 3.1415))
	s.Close()

	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet := xf.Sheets[0]
	assert.Equal(t,
		[]streamxlsx.TestCell{
			{"A1", "inlineStr", "a float", 0},
			{"B1", "n", "3.141500", 0},
			{"A2", "inlineStr", "a styled float (default)", 0},
			{"B2", "n", "3.141500", 1},
			{"A3", "inlineStr", "a styled float (custom)", 0},
			{"B3", "n", "3.141500", 2},
		},
		sheet.Cells,
	)
}

func TestHangingSheet(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteRow("aap")
	s.WriteRow("noot")
	// no s.WriteSheet
	s.Close()

	// Read it back again
	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Cells, 2)
}

func TestEmptyFile(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.Close()

	// Read it back again
	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 1)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Cells, 0)
}

func TestEmptySheet(t *testing.T) {
	buf := &bytes.Buffer{}
	s := streamxlsx.New(buf)
	s.WriteSheet("sheet 1")
	s.WriteSheet("sheet 2")
	s.Close()

	// Read it back again
	xf, err := streamxlsx.TestParse(buf.Bytes())
	require.NoError(t, err)
	require.Len(t, xf.Sheets, 2)
	sheet0 := xf.Sheets[0]
	require.Equal(t, "sheet 1", sheet0.Name)
	require.Len(t, sheet0.Cells, 0)
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
