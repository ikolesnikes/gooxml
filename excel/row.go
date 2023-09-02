package excel

import (
	"encoding/xml"
	"fmt"

	"golang.org/x/exp/slices"
)

type row struct {
	index int
	cells map[int]*cell
}

func newRow(i int) *row {
	r := row{
		index: i,
		cells: make(map[int]*cell),
	}
	return &r
}

func (r *row) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
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
