package excel

import "testing"

func Test_MakeContentTypesNewDoc(t *testing.T) {
	doc := NewDocument()
	cts := makeContentTypes(doc)
	wantTyps := map[string]struct{}{
		"rels":                      {},
		"xml":                       {},
		"/xl/workbook.xml":          {},
		"/xl/sharedStrings.xml":     {},
		"/xl/worksheets/sheet1.xml": {},
	}
	if len(cts.items) != len(wantTyps) {
		t.Errorf("Want %d, got %d", len(wantTyps), len(cts.items))
	}
	for _, ct := range cts.items {
		_, ok := wantTyps[ct.content()]
		if !ok {
			t.Errorf("Type %q not found", ct.content())
		}
	}
}
