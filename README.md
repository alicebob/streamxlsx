streaming XLSX writer for simple tabular data

An .xlsx file is not much more than some gzipped XML files. If all you want is to generate some spreasheets with tabular data, there is no reason we can't generate the file completely streaming.

This small library does just that. It's a small API to generate an XLSX file on the fly, streaming. The main goal is to have a simple way to offer exports of data, for example via http.

See `./example_test.go` how to use this.

Example:
```
	s := streamxls.New(buf)
	s.WriteRow("first row with a simple string", 3.1415)
	s.WriteRow("this is row 2")
	s.WriteRow("3digits pi:", s.Format("0.000", 3.1415))
	s.WriteRow("click there:", Hyperlink{"http://example.com", "clickme", "I'm a tooltip"})
	// close the open sheet, with this title. The next WriteRow() goes into a new sheet
	s.WriteSheet("that was sheet 1")
	s.WriteRow("13")
	s.Close()
```

## features

- streams (almost) the whole file
- support for number formatting
- hyperlinks
- multiple sheets
- currently no support for colors, fonts, borders, graphs, or formulas.


## status

"seems to work". The files have been tested on: gnumeric, Goog spreadsheets, online office 365 spreadsheet.
The exact API how to style cells will still change.


## see also

https://github.com/tealeg/xlsx/
https://docs.microsoft.com/en-us/office/open-xml/working-with-sheets
https://github.com/psmithuk/xlsx
https://github.com/plandem/xlsx
