package excel

import (
	"encoding/xml"
	"slices"
)

// A worksheet inside Excel's workbook.
type Worksheet struct {
	id     int
	wkbRel *relationship
	rows   map[int]*Row
}

// newWorksheet creates and initializes a new worksheet.
func newWorksheet(id int, rel *relationship) *Worksheet {
	wks := Worksheet{
		id:     id,
		wkbRel: rel,
		rows:   make(map[int]*Row),
	}
	return &wks
}

// SetText writes a string value into the cell.
// Row and cell indices are zero-based.
func (wks *Worksheet) SetText(t string, ri int, ci int) {
	r := wks.rows[ri]
	if r == nil {
		r = newRow(ri)
		wks.rows[ri] = r
	}

	c := r.cells[ci]
	if c == nil {
		c = newCell(ri, ci)
		r.cells[ci] = c
	}

	c.SetText(t)
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

	// Write rows sorted out by index

	// Get keys and sort them
	keys := make([]int, len(wks.rows))
	i := 0
	for k := range wks.rows {
		keys[i] = k
		i++
	}
	slices.Sort(keys)

	for _, i := range keys {
		r := wks.rows[i]
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
