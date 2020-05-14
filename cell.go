package streamxlsx

import (
	"bufio"
	"encoding/base64"
	"encoding/xml"
	"fmt"
	"strconv"
	"time"
)

// These can be passed to `WriteRow()` if you want total control. WriteRow()
// will fill int the `.Ref` value.
// Exactly one of Value or InlineString should be set.
type Cell struct {
	Ref          string  `xml:"r,attr"` // "A1" &c.
	Type         string  `xml:"t,attr,omitempty"`
	Style        *int    `xml:"s,attr,omitempty"`
	Value        string  `xml:"v,omitempty"`
	InlineString *string `xml:"is>t,omitempty"`
	hyperlink    *hyperlink
}

func writeCell(w *bufio.Writer, c Cell) error {
	w.WriteString(`<c r="`)
	w.WriteString(c.Ref) // no enc
	if c.Type != "" {
		w.WriteString(`" t="`)
		w.WriteString(c.Type) // no enc
	}
	if c.Style != nil {
		w.WriteString(`" s="`)
		w.WriteString(strconv.Itoa(*c.Style))
	}
	w.WriteString(`">`)
	if c.InlineString != nil {
		w.WriteString(`<is><t>`)
		xml.EscapeText(w, []byte(*c.InlineString))
		w.WriteString(`</t></is>`)
	} else {
		w.WriteString(`<v>`)
		xml.EscapeText(w, []byte(c.Value))
		w.WriteString(`</v>`)
	}
	w.WriteString(`</c>`)
	return nil
}

func (c Cell) String() string {
	if c.InlineString != nil {
		return *c.InlineString
	}
	return c.Value
}

func asCell(v interface{}) (Cell, error) {
	switch vt := v.(type) {
	case Cell:
		return vt, nil
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%d", vt),
		}, nil
	case float32, float64:
		return Cell{
			Type:  "n",
			Value: fmt.Sprintf("%f", vt),
		}, nil
	case []byte:
		return asCell(base64.StdEncoding.EncodeToString(vt))
	case string:
		return Cell{
			Type:         "inlineStr",
			InlineString: &vt,
		}, nil
	case time.Time:
		return Cell{
			Type:  "n",
			Value: oaDate(vt),
		}, nil
	case bool:
		v := "0"
		if vt {
			v = "1"
		}
		return Cell{
			Type:  "b",
			Value: v,
		}, nil
	case Hyperlink:
		cell, err := asCell(vt.Title)
		cell.hyperlink = &hyperlink{
			url:     vt.URL,
			Display: vt.Title,
			Tooltip: vt.Tooltip,
		}
		return cell, err
	default:
		return Cell{}, fmt.Errorf("unsupported cell type: %T", vt)
	}
}

func applyStyle(id int, v interface{}) (Cell, error) {
	c, err := asCell(v)
	c.Style = &id
	return c, err
}

func oaDate(d time.Time) string {
	// taken from https://github.com/psmithuk/xlsx/blob/master/xlsx.go
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	nsPerDay := 24 * time.Hour

	// keep times in the given timezone
	fakeUTC := time.Date(d.Year(), d.Month(), d.Day(), d.Hour(), d.Minute(), d.Second(), d.Nanosecond(), time.UTC)

	v := -1 * float64(epoch.Sub(fakeUTC)) / float64(nsPerDay)

	// TODO: deal with dates before epoch
	// e.g. http://stackoverflow.com/questions/15549823/oadate-to-milliseconds-timestamp-in-javascript/15550284#15550284

	if fakeUTC.Hour() == 0 && fakeUTC.Minute() == 0 && fakeUTC.Second() == 0 {
		return fmt.Sprintf("%d", int64(v))
	}
	return fmt.Sprintf("%f", v)
}
