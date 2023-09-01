package excel

import (
	"encoding/xml"
	"fmt"
)

// Shared strings part.
type sharedStrings struct {
	// Strings table. Key - string, value - count of occurences of the string.
	st map[string]int
}

// newSharedStrings creates and initializes a new shared strings item.
func newSharedStrings() *sharedStrings {
	return &sharedStrings{
		st: make(map[string]int),
	}
}

func (sst *sharedStrings) add(s string) {
	sst.st[s]++
}

// count counts total and unique strings in the table.
func (sst *sharedStrings) count() (total int, unique int) {
	for _, v := range sst.st {
		total += v
		if v == 1 {
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

	for str := range sst.st {
		siName := xml.Name{Local: "si"}
		tName := xml.Name{Local: "t"}
		tokens := []xml.Token{
			xml.StartElement{Name: siName},
			xml.StartElement{Name: tName},
			str,
			xml.EndElement{Name: tName},
			xml.EndElement{Name: siName},
		}
		if err := encodeTokens(tokens, enc); err != nil {
			return err
		}
	}

	return enc.EncodeToken(xml.EndElement{Name: sstName})
}
