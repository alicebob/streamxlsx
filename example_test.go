package streamxlsx

import (
	"os"
)

func Example() {
	fh, err := os.Create("./example.xlsx")
	if err != nil {
		panic(err)
	}
	defer fh.Close()

	s := New(fh)
	// s.WriteRow("foo", 12, time.Now(), streamxlsx.Cell(12.3), CellFloatWithformat(12.3, "0.00"))
	s.WriteRow("12")
	s.WriteSheet("that was sheet 1")

	s.WriteRow("13")
	s.WriteSheet("that was sheet 2")
	s.Close()

	// Output:
}
