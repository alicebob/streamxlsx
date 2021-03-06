package streamxlsx

import (
	"os"
	"time"
)

func Example() {
	fh, err := os.Create("./example.xlsx")
	if err != nil {
		panic(err)
	}
	defer fh.Close()

	s := New(fh)
	s.WriteRow("a string", "hello world!", "expected: 'hello world!'")
	s.WriteRow("also a string", "14", "expected: '14'")
	s.WriteRow("a number", 14, "expected: 14")
	s.WriteRow("a number negative number", s.Format("#,##0 ;[red](#,##0)", -14), "expected: (14)-red")
	s.WriteRow("a float", 3.1415, "expected: 3.1415")
	s.WriteRow("a float, formatted", s.Format("0.00", 3.1415), "expected: 3.14")
	s.WriteRow("a float, also formatted", s.Format("0.000", 3.1415), "expected: 3.142")
	s.WriteRow("a link", Hyperlink{"http://example.com", "clickme", "I'm a tooltip"}, "expected: link to 'http://example.com', title'clickme', tooltip 'I'm a tooltip'")
	s.WriteRow("a link", Hyperlink{"http://example.com/v2", "clickmev2", "I'm a tooltipv2"}, "expected: link to 'http://example.com/v2', title'clickmev2', tooltip 'I'm a tooltipv2'")
	s.WriteRow("a datetime", s.Format(DefaultDatetimeFormat, time.Date(2010, 10, 10, 10, 10, 10, 0, time.UTC)), "expected: 10/10/2010 10:10 (or 10/10/10)")
	s.WriteRow("bools", true, false)
	s.WriteRow("empty cell", nil, "<-- empty cell")
	s.WriteRow()
	s.WriteRow()
	s.WriteRow("there should be another sheet with a single value")
	s.WriteSheet("that was sheet 1")

	s.WriteRow("13")
	s.WriteSheet("that was sheet 2")
	if err := s.Close(); err != nil {
		panic(err)
	}

	// Output:
}
