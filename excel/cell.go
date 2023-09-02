package excel

import (
	"encoding/xml"
)

// A cell in the worksheet.
type cell struct {
	// Text contained in the cell.
	// If shared strings, this is an index into shared strings table.
	text string

	// Row and column indexes of this cell.
	ri, ci int
}

// newCell creates and initializes a new cell structure.
func newCell(ri, ci int) *cell {
	c := cell{
		ri: ri,
		ci: ci,
	}
	return &c
}

func (c *cell) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	cName := xml.Name{Local: "c"}

	cStart := xml.StartElement{
		Name: cName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "r"}, Value: makeRef(c.ri, c.ci)},
			{Name: xml.Name{Local: "t"}, Value: "s"},
		},
	}

	tokens := []xml.Token{
		cStart,
		xml.StartElement{Name: xml.Name{Local: "v"}},
		xml.CharData(c.text),
		xml.EndElement{Name: xml.Name{Local: "v"}},
		xml.EndElement{Name: cName},
	}
	return encodeTokens(tokens, enc)
}
