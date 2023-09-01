package excel

// Content-Types part

import "encoding/xml"

// Well known content types.
const (
	CTRels      = "application/vnd.openxmlformats-package.relationships+xml"
	CTXML       = "application/xml"
	CTWorkbook  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	CTWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)

// A collection of content-type items.
type contentTypes struct {
	ctyps []*contentTypeConformant
}

// newContentTypes creates and initializes a new content-types part.
func newContentTypes() *contentTypes {
	cts := contentTypes{}
	cts.add(newContentTypeDefault("rels", CTRels))
	cts.add(newContentTypeDefault("xml", CTXML))
	return &cts
}

func (cts *contentTypes) add(ct contentTypeConformant) {
	cts.ctyps = append(cts.ctyps, &ct)
}

func (cts *contentTypes) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: "Types"}

	start := xml.StartElement{
		Name: name,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSPackageCT},
		},
	}

	tokens := []xml.Token{
		xmlProlog,
		start,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	if err := enc.EncodeElement(cts.ctyps, start); err != nil {
		return err
	}

	if err := enc.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}

type contentTypeConformant interface {
	content() string
	typ() string
}

// Base type for different content types. There are two known content types:
// Default and Override.
type contentType struct {
	c string
	t string
}

// newContentType creates and initializes a new contentType.
func newContentType(c, t string) *contentType {
	ct := contentType{c, t}
	return &ct
}

func (ct *contentType) content() string {
	return ct.c
}

func (ct *contentType) typ() string {
	return ct.t
}

// toXML marshals the content type to XML. This is a common method for
// different content types which only differ in element/attribute names.
func (ct *contentType) toXML(eName string, aName string, enc *xml.Encoder, root xml.StartElement) error {
	name := xml.Name{Local: eName}
	tokens := []xml.Token{
		xml.StartElement{
			Name: name,
			Attr: []xml.Attr{
				{Name: xml.Name{Local: aName}, Value: ct.content()},
				{Name: xml.Name{Local: "ContentType"}, Value: ct.typ()},
			},
		},
		xml.EndElement{Name: name},
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}
	return nil
}

type contentTypeDefault struct {
	*contentType
}

func newContentTypeDefault(c, t string) *contentTypeDefault {
	ct := contentTypeDefault{
		contentType: newContentType(c, t),
	}
	return &ct
}

// <Default Extension="{{C}}" ContentType="{{T}}"/>
func (ct *contentTypeDefault) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	return ct.toXML("Default", "Extension", enc, root)
}

type contentTypeOverride struct {
	*contentType
}

func newContentTypeOverride(c, t string) *contentTypeOverride {
	ct := contentTypeOverride{
		contentType: newContentType(c, t),
	}
	return &ct
}

// <Override PartName="{{C}}" ContentType="{{T}}"/>
func (ct *contentTypeOverride) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	return ct.toXML("Override", "PartName", enc, root)
}
