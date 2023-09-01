package excel

import "encoding/xml"

// A worksheet
type Worksheet struct {
	id  int
	rel *relationship
}

// newWorksheet creates and initializes a new worksheet.
func newWorksheet(id int, rel *relationship) *Worksheet {
	wks := Worksheet{
		id:  id,
		rel: rel,
	}
	return &wks
}

func (wks *Worksheet) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: "worksheet"}

	start := xml.StartElement{
		Name: name,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSSpreadsheetML},
		},
	}

	tokens := []xml.Token{
		xmlProlog,
		start,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	e := xml.StartElement{Name: xml.Name{Local: "sheetData"}}
	if err := enc.EncodeToken(e); err != nil {
		return err
	}
	if err := enc.EncodeToken(xml.EndElement{Name: xml.Name{Local: "sheetData"}}); err != nil {
		return err
	}
	if err := enc.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}
