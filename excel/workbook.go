package excel

import (
	"encoding/xml"
	"fmt"
)

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
	// New worksheet's id and file name
	id := wkb.newSheetID()
	fName := fmt.Sprintf("sheet%d.xml", id)

	// Workbook -> Worksheet relationship item
	rel := newRelationship(wkb.rels.newID(), RELWorksheet, fmt.Sprintf("worksheets/%s", fName))
	wkb.rels.add(rel)

	// Add new worksheet to the collection
	wkb.sheets = append(wkb.sheets, newWorksheet(id, rel))

	// Worksheet's content-type entry
	wkb.doc.cts.add(newContentTypeOverride(fmt.Sprintf("/xl/worksheets/%s", fName), CTWorksheet))
}

func (wkb *Workbook) newSheetID() int {
	id := len(wkb.sheets) + 1
	return id
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
	for _, wks := range wkb.sheets {
		e := xml.StartElement{
			Name: xml.Name{Local: "sheet"},
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "name"}, Value: fmt.Sprintf("Sheet%d", wks.id)},
				{Name: xml.Name{Local: "sheetId"}, Value: fmt.Sprintf("%d", wks.id)},
				{Name: xml.Name{Local: "r:id"}, Value: wks.rel.id},
			},
		}
		if err := enc.EncodeToken(e); err != nil {
			return err
		}
		if err := enc.EncodeToken(xml.EndElement{Name: xml.Name{Local: "sheet"}}); err != nil {
			return err
		}
	}
	if err := enc.EncodeToken(xml.EndElement{Name: xml.Name{Local: "sheets"}}); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}
