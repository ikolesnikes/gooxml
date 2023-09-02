package excel

import (
	"encoding/xml"
)

type cell struct {
	text string
}

func newCell() *cell {
	return &cell{}
}

func (c *cell) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	cName := xml.Name{Local: "c"}

	cStart := xml.StartElement{
		Name: cName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "r"}, Value: "A1"},
			{Name: xml.Name{Local: "t"}, Value: "s"},
		},
	}

	tokens := []xml.Token{
		cStart,
		xml.StartElement{Name: xml.Name{Local: "v"}},
		xml.CharData("0"),
		xml.EndElement{Name: xml.Name{Local: "v"}},
		xml.EndElement{Name: cName},
	}
	return encodeTokens(tokens, enc)
}
