package excel

import "encoding/xml"

type Worksheet struct {
}

func newWorksheet() *Worksheet {
	wks := Worksheet{}
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
