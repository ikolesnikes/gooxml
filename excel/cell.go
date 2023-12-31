package excel

import (
	"encoding/xml"
	"strconv"
)

// A Cell in the worksheet.
type Cell struct {
	// Text contained in the cell.
	text string

	// Cell's text index in the shared strings.
	// This is being set during document save.
	// After the save it's not guaranteed to be valid.
	sstIndex int

	// Row and column indices of this cell.
	ri, ci int
}

// newCell creates and initializes a new cell structure.
func newCell(ri, ci int) *Cell {
	c := Cell{
		ri: ri,
		ci: ci,
	}
	return &c
}

// SetText writes text t to the cell.
func (c *Cell) SetText(t string) {
	c.text = t
}

func (c *Cell) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	cName := xml.Name{Local: "c"}

	tokens := []xml.Token{
		xml.StartElement{
			Name: cName,
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "r"}, Value: makeA1Ref(c.ri, c.ci)},
				{Name: xml.Name{Local: "t"}, Value: "s"},
			},
		},
		xml.StartElement{Name: xml.Name{Local: "v"}},
		xml.CharData(strconv.Itoa(c.sstIndex)),
		xml.EndElement{Name: xml.Name{Local: "v"}},
		xml.EndElement{Name: cName},
	}
	return encodeTokens(tokens, enc)
}
