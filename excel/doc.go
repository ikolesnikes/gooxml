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
	cts  *contentTypes
	rels *relationships
	wkb  *Workbook
}

// NewDocument creates and initializes a new Excel document. It returns
// a document containing workbook and a single worksheet. This resambles
// the case of running Excel program and creating a new blank document.
// Saving this document produces a valid .xlsx file.
func NewDocument() *Document {
	doc := Document{
		cts:  newContentTypes(),
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
	doc.cts.add(newContentTypeOverride("/xl/workbook.xml", CTWorkbook))
}

// Save writes the document using the given writer. The written document is the
// complete (and hopefully valid) .xlsx file.
func (doc *Document) Save(w io.Writer) error {

	// Find all parts that need to be encoded and encode them all here
	// or...
	// pass the 'save' command down each level...?
	//
	// At the document level (this, top-most level) there are
	// content-types
	// package relationships
	// workbook
	//
	// At the workbook level there are
	// workbook relationships
	// worksheets

	type partDesc struct {
		path string
		part xml.Marshaler
		body *bytes.Buffer
	}
	var parts = []*partDesc{
		{"[Content_Types].xml", doc.cts, nil},
		{"_rels/.rels", doc.rels, nil},
		{"xl/workbook.xml", doc.wkb, nil},
		{"xl/_rels/workbook.xml.rels", doc.wkb.rels, nil},
	}
	for _, wks := range doc.wkb.sheets {
		parts = append(parts, &partDesc{fmt.Sprintf("xl/worksheets/sheet%d.xml", wks.id), wks, nil})
	}

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
