package streamxlsx_test

import (
	"bytes"
	"log"
	"os"
	"runtime/pprof"
	"strconv"
	"testing"

	"github.com/alicebob/streamxlsx"
	"github.com/tealeg/xlsx/v2"
)

func BenchmarkStream(b *testing.B) {
	if true {
		f, err := os.Create("/tmp/stream.pprof")
		if err != nil {
			log.Fatal("could not create CPU profile: ", err)
		}
		defer f.Close() // error handling omitted for example
		if err := pprof.StartCPUProfile(f); err != nil {
			log.Fatal("could not start CPU profile: ", err)
		}
		defer pprof.StopCPUProfile()
	}

	for i := 0; i < b.N; i++ {
		buf := &bytes.Buffer{}
		s := streamxlsx.New(buf)
		for j := 0; j < 1_000_000; j++ {
			s.WriteRow(j, s.Format("0.00", j), strconv.Itoa(j))
		}
		s.Close()
	}
}

func BenchmarkTealeg(b *testing.B) {
	for i := 0; i < b.N; i++ {
		f := xlsx.NewFile()
		sheet, _ := f.AddSheet("foo")
		for j := 0; j < 1_000_000; j++ {
			row := sheet.AddRow()
			row.AddCell().SetInt(j)
			cell := row.AddCell()
			cell.SetInt(j)
			cell.SetFormat("0.00")
			row.AddCell().SetValue(strconv.Itoa(j))
		}
		buf := &bytes.Buffer{}
		f.Write(buf)
	}
}
