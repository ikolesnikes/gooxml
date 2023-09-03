package excel

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
)

// A Document represents Excel's document package.
type Document struct {
	rels *relationships
	wkb  *Workbook
}

// NewDocument creates and initializes a new Excel document. It returns
// a document containing workbook and a single worksheet. This resambles
// the case of running Excel program and creating a new blank document.
// Saving this document produces a valid .xlsx file.
func NewDocument() *Document {
	doc := Document{
		rels: newRelationships(),
	}
	doc.addWorkbook()
	return &doc
}

// Workbook returns the document's workbook.
func (doc *Document) Workbook() *Workbook {
	return doc.wkb
}

func (doc *Document) addWorkbook() {
	doc.wkb = newWorkbook(doc)
	doc.rels.add(newRelationship(doc.rels.newID(), RELOfficeDocument, "xl/workbook.xml"))
}

// Save writes the document using the given writer. The written document is the
// complete (and hopefully valid) .xlsx file.
func (doc *Document) Save(w io.Writer) error {
	type partDesc struct {
		path string
		part xml.Marshaler
		body *bytes.Buffer
	}
	var parts = []*partDesc{
		{path: "_rels/.rels", part: doc.rels},
		{path: "xl/workbook.xml", part: doc.wkb},
		{path: "xl/sharedStrings.xml", part: doc.wkb.sst},
		{path: "xl/_rels/workbook.xml.rels", part: doc.wkb.rels},
	}
	for _, wks := range doc.wkb.sheets {
		parts = append(parts, &partDesc{path: fmt.Sprintf("xl/worksheets/sheet%d.xml", wks.id), part: wks})
	}
	parts = append(parts, &partDesc{path: "[Content_Types].xml", part: makeContentTypes(doc)})

	// This can be done in parallel
	var err error
	for _, part := range parts {
		part.body, err = encodePart(part.part)
		if err != nil {
			return err
		}
	}

	z := zip.NewWriter(w)
	for _, part := range parts {
		f, err := z.Create(part.path)
		if err != nil {
			return err
		}
		_, err = f.Write(part.body.Bytes())
		if err != nil {
			return err
		}
	}
	return z.Close()
}

// makeContentTypes creates the content types part. The part returned
// is populated with all needed entries based on the given document.
func makeContentTypes(doc *Document) *contentTypes {
	cts := newContentTypes()
	cts.add(newContentTypeOverride("/xl/workbook.xml", CTWorkbook))
	cts.add(newContentTypeOverride("/xl/sharedStrings.xml", CTSharedStrings))
	for _, wks := range doc.wkb.sheets {
		fName := fmt.Sprintf("sheet%d.xml", wks.id)
		cts.add(newContentTypeOverride(fmt.Sprintf("/xl/worksheets/%s", fName), CTWorksheet))
	}
	return cts
}
