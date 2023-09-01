package excel

import "encoding/xml"

// A workbook inside an Excel document.
type Workbook struct {
	doc    *Document
	rels   *relationships
	sheets []*Worksheet
}

// newWorkbook creates and initializes a new Workbook.
func newWorkbook(doc *Document) *Workbook {
	wkb := Workbook{
		doc:  doc,
		rels: newRelationships(),
	}
	wkb.AddWorksheet()
	return &wkb
}

func (wkb *Workbook) AddWorksheet() {
	wks := newWorksheet()
	wkb.sheets = append(wkb.sheets, wks)
	wkb.rels.add(newRelationship(wkb.rels.newID(), RELWorksheet, "worksheets/sheet1.xml"))
	wkb.doc.cts.add(newContentTypeOverride("/xl/worksheets/sheet1.xml", CTWorksheet))
}

func (wkb *Workbook) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: "workbook"}

	start := xml.StartElement{
		Name: name,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSSpreadsheetML},
			{Name: xml.Name{Local: "xmlns:r"}, Value: NSOfficeDocRels},
		},
	}

	tokens := []xml.Token{
		xmlProlog,
		start,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.StartElement{Name: xml.Name{Local: "sheets"}}); err != nil {
		return err
	}
	// for _, s := range w.sheets {
	e := xml.StartElement{
		Name: xml.Name{Local: "sheet"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "name"}, Value: "Sheet1"},
			{Name: xml.Name{Local: "sheetId"}, Value: "1"},
			{Name: xml.Name{Local: "r:id"}, Value: "rId1"},
		},
	}
	if err := enc.EncodeToken(e); err != nil {
		return err
	}
	if err := enc.EncodeToken(xml.EndElement{Name: xml.Name{Local: "sheet"}}); err != nil {
		return err
	}
	// }
	if err := enc.EncodeToken(xml.EndElement{Name: xml.Name{Local: "sheets"}}); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}
