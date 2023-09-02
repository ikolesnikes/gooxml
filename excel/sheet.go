package excel

import (
	"encoding/xml"
	"fmt"
)

// A worksheet
type Worksheet struct {
	id     int
	wkbRel *relationship
	wkb    *Workbook
	rows   map[int]*row
}

// newWorksheet creates and initializes a new worksheet.
func newWorksheet(id int, rel *relationship, wkb *Workbook) *Worksheet {
	wks := Worksheet{
		id:     id,
		wkbRel: rel,
		wkb:    wkb,
		rows:   make(map[int]*row),
	}
	return &wks
}

func (wks *Worksheet) AddText(s string, ri int, ci int) {
	r := wks.rows[ri]
	if r == nil {
		r = newRow(ri)
		wks.rows[ri] = r
	}

	c := r.cells[ci]
	if c == nil {
		c = newCell()
		r.cells[ci] = c
	}

	i := wks.wkb.sst.add(s)
	c.text = fmt.Sprintf("%d", i)
}

func (wks *Worksheet) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	worksheetName := xml.Name{Local: "worksheet"}
	sheetDataName := xml.Name{Local: "sheetData"}

	worksheetStart := xml.StartElement{
		Name: worksheetName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSSpreadsheetML},
		},
	}

	sheetDataStart := xml.StartElement{Name: sheetDataName}

	tokens := []xml.Token{
		xmlProlog,
		worksheetStart,
		sheetDataStart,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	for _, r := range wks.rows {
		if err := enc.EncodeElement(r, sheetDataStart); err != nil {
			return err
		}
	}

	tokens = []xml.Token{
		xml.EndElement{Name: sheetDataName},
		xml.EndElement{Name: worksheetName},
	}
	return encodeTokens(tokens, enc)
}
