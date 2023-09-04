package excel

import (
	"encoding/xml"
	"fmt"
	"slices"
)

// A row inside Excel's worksheet.
type Row struct {
	index int
	cells map[int]*Cell
}

func newRow(i int) *Row {
	r := Row{
		index: i,
		cells: make(map[int]*Cell),
	}
	return &r
}

func (r *Row) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	rowName := xml.Name{Local: "row"}

	rowStart := xml.StartElement{
		Name: rowName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "r"}, Value: fmt.Sprintf("%d", r.index+1)},
			{Name: xml.Name{Local: "spans"}, Value: "1:1"},
		},
	}

	if err := enc.EncodeToken(rowStart); err != nil {
		return err
	}

	// Write cells sorted out by index

	// Get keys and sort them
	keys := make([]int, len(r.cells))
	i := 0
	for k := range r.cells {
		keys[i] = k
		i++
	}
	slices.Sort(keys)

	for _, i := range keys {
		c := r.cells[i]
		if err := enc.EncodeElement(c, rowStart); err != nil {
			return err
		}
	}

	return enc.EncodeToken(xml.EndElement{Name: rowName})
}
