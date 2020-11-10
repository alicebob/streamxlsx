streaming XLSX writer for simple tabular data

An .xlsx file is not much more than some gzipped XML files. If all you want is to generate some spreadsheets with tabular data, there is no reason we can't generate the file completely streaming.

This small library does just that. It generates an XLSX file on the fly, streaming. The main goal is to have a way to offer exports of data, for example via http.


## Example:
See also `./example_test.go`.

```
	s := streamxls.New(buf)
	s.WriteRow("first row with a simple string", 3.1415)
	s.WriteRow("this is row 2")
	s.WriteRow("3digits pi:", s.Format("0.000", 3.1415))
	s.WriteRow("click there:", Hyperlink{"http://example.com", "clickme", "I'm a tooltip"})
	s.WriteSheet("that was sheet 1")
	s.WriteRow("13") // that's a new sheet
	s.Close()
```

## features

- streams (almost) the whole file
- support for basic spreadsheet features: number formatting, hyperlinks, sheets
- currently no support for colors, fonts, borders
- likely never support for graphs, merged cells, hidden columns, or formulas.


## status

"seems to work". The files have been tested with: gnumeric, Goog spreadsheets, online office 365 Excel, offline Excel, Numbers, Emacs, LibreOffice Calc.


## see also
https://github.com/TheDataShed/xlsxreader  
https://github.com/tealeg/xlsx/  
https://docs.microsoft.com/en-us/office/open-xml/working-with-sheets  
https://github.com/psmithuk/xlsx  

Limits for xlsx files:
https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
