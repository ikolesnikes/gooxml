package excel

import (
	"encoding/xml"
	"fmt"
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

	for _, c := range r.cells {
		if err := enc.EncodeElement(c, rowStart); err != nil {
			return err
		}
	}

	return enc.EncodeToken(xml.EndElement{Name: rowName})
}
