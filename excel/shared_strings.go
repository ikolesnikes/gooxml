package excel

import (
	"cmp"
	"encoding/xml"
	"fmt"
	"slices"
)

// Shared strings table.
type sharedStrings struct {
	items map[string]*stringItem
}

// Item in the shared strings table.
type stringItem struct {
	text string

	// Number of occurences of this item in the table.
	// The equal items aren't duplicated.
	count int

	// Zero-based index of this item into the table.
	index int
}

// newSharedStrings creates and initializes a new shared strings table.
func newSharedStrings() *sharedStrings {
	return &sharedStrings{
		items: make(map[string]*stringItem),
	}
}

// add adds a new string item, containing the text, into the table.
// Returns the item's index.
func (sst *sharedStrings) add(text string) int {
	si := sst.items[text]
	if si == nil {
		si = &stringItem{
			text:  text,
			index: len(sst.items),
		}
		sst.items[text] = si
	}
	si.count++
	return si.index
}

// count counts total and unique strings in the table.
func (sst *sharedStrings) count() (total int, unique int) {
	for _, si := range sst.items {
		total += si.count
		if si.count == 1 {
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

	// Sort items by index before encoding
	// TODO: better way for managing string items: order by index, fast access
	ordered := make([]*stringItem, len(sst.items))
	i := 0
	for _, si := range sst.items {
		ordered[i] = si
		i++
	}
	slices.SortFunc(ordered, func(a, b *stringItem) int {
		return cmp.Compare(a.index, b.index)
	})

	for _, si := range ordered {
		siName := xml.Name{Local: "si"}
		tName := xml.Name{Local: "t"}
		tokens := []xml.Token{
			xml.StartElement{Name: siName},
			xml.StartElement{Name: tName},
			xml.CharData(si.text),
			xml.EndElement{Name: tName},
			xml.EndElement{Name: siName},
		}
		if err := encodeTokens(tokens, enc); err != nil {
			return err
		}
	}

	return enc.EncodeToken(xml.EndElement{Name: sstName})
}
