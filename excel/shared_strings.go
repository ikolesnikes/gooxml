package excel

import (
	"encoding/xml"
	"fmt"
)

// Shared strings part.
type sharedStrings struct {
	st []stringEntry
}

type stringEntry struct {
	s string
	c int
}

// newSharedStrings creates and initializes a new shared strings item.
func newSharedStrings() *sharedStrings {
	return &sharedStrings{}
}

func (sst *sharedStrings) add(s string) int {
	// Find string 's' in table
	// Find string's 's' index in table

	var i int
	for i = 0; i < len(sst.st); i++ {
		if s == sst.st[i].s {
			break
		}
	}
	if i == len(sst.st) {
		// String 's' not found
		sst.st = append(sst.st, stringEntry{s: s})
		i = len(sst.st) - 1
	}

	sst.st[i].s = s
	sst.st[i].c++
	return i
}

// count counts total and unique strings in the table.
func (sst *sharedStrings) count() (total int, unique int) {
	for _, e := range sst.st {
		total += e.c
		if e.c == 1 {
			unique++
		}
	}
	return
}

func (sst *sharedStrings) MarshalXML(enc *xml.Encoder, root xml.StartElement) error {
	total, unique := sst.count()

	sstName := xml.Name{Local: "sst"}

	start := xml.StartElement{
		Name: sstName,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: NSSpreadsheetML},
			{Name: xml.Name{Local: "count"}, Value: fmt.Sprintf("%d", total)},
			{Name: xml.Name{Local: "uniqueCount"}, Value: fmt.Sprintf("%d", unique)},
		},
	}
	tokens := []xml.Token{
		xmlProlog,
		start,
	}
	if err := encodeTokens(tokens, enc); err != nil {
		return err
	}

	for _, e := range sst.st {
		siName := xml.Name{Local: "si"}
		tName := xml.Name{Local: "t"}
		tokens := []xml.Token{
			xml.StartElement{Name: siName},
			xml.StartElement{Name: tName},
			xml.CharData(e.s),
			xml.EndElement{Name: tName},
			xml.EndElement{Name: siName},
		}
		if err := encodeTokens(tokens, enc); err != nil {
			return err
		}
	}

	return enc.EncodeToken(xml.EndElement{Name: sstName})
}
