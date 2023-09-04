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
	wkb.addSharedStrings()
	wkb.AddWorksheet()
	return &wkb
}

func (wkb *Workbook) addSharedStrings() {
	rel := newRelationship(wkb.rels.newID(), RELSharedStrings, "sharedStrings.xml")
	wkb.rels.add(rel)
}

// Worksheet returns a worksheet by its id.
func (wkb *Workbook) Worksheet(id int) *Worksheet {
	return wkb.sheets[id]
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
}

func (wkb *Workbook) newSheetID() int {
	id := len(wkb.sheets) + 1
	return id
}

func (wkb *Workbook) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	workbookName := xml.Name{Local: "workbook"}
	sheetsName := xml.Name{Local: "sheets"}
	start := xml.StartElement{
		Name: workbookName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSSpreadsheetML},
			{Name: xml.Name{Local: "xmlns:r"}, Value: NSOfficeDocRels},
		},
	}
	tokens := []xml.Token{
		xmlProlog,
		start,
		xml.StartElement{Name: sheetsName},
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	sheetName := xml.Name{Local: "sheet"}
	for _, wks := range wkb.sheets {
		start := xml.StartElement{
			Name: sheetName,
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "name"}, Value: fmt.Sprintf("Sheet%d", wks.id)},
				{Name: xml.Name{Local: "sheetId"}, Value: fmt.Sprintf("%d", wks.id)},
				{Name: xml.Name{Local: "r:id"}, Value: wks.wkbRel.id},
			},
		}
		tokens := []xml.Token{
			start,
			xml.EndElement{Name: sheetName},
		}
		if err := encodeTokens(tokens, enc); err != nil {
			return err
		}
	}

	tokens = []xml.Token{
		xml.EndElement{Name: sheetsName},
		xml.EndElement{Name: workbookName},
	}
	return encodeTokens(tokens, enc)
}
