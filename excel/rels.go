package excel

import (
	"encoding/xml"
	"fmt"
)

const (
	RELOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RELWorksheet      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	RELSharedStrings  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)

// A relationships part.
type relationships struct {
	rs []*relationship
}

// newRelationships creates and initializes a new relationships part.
func newRelationships() *relationships {
	return &relationships{}
}

// newID returns a new id for an relationship. All ids within the part
// must be unique.
func (rels *relationships) newID() string {
	id := len(rels.rs) + 1
	return fmt.Sprintf("rId%d", id)
}

func (rels *relationships) add(rel *relationship) {
	rels.rs = append(rels.rs, rel)
}

func (rels *relationships) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: "Relationships"}

	start := xml.StartElement{
		Name: name,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSPackageRels},
		},
	}

	tokens := []xml.Token{
		xmlProlog,
		start,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	if err := enc.EncodeElement(rels.rs, start); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}

// Single relationship definition.
type relationship struct {
	id, typ, target string
}

func newRelationship(id, typ, target string) *relationship {
	return &relationship{id: id, typ: typ, target: target}
}

func (rel *relationship) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: "Relationship"}
	tokens := []xml.Token{
		xml.StartElement{
			Name: name,
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "Id"}, Value: rel.id},
				{Name: xml.Name{Local: "Type"}, Value: rel.typ},
				{Name: xml.Name{Local: "Target"}, Value: rel.target},
			},
		},
		xml.EndElement{Name: name},
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}
	return nil
}
