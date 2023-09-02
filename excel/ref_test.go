package excel

import "testing"

func Test_indexToColumnName(t *testing.T) {
	table := []struct {
		i    int
		want string
	}{
		{0, "A"},
		{1, "B"},
		{25, "Z"},
		{26, "AA"},
		{27, "AB"},
		{51, "AZ"},
		{52, "BA"},
		{53, "BB"},
		{701, "ZZ"},
		{702, "AAA"},
		{15465, "VVV"},
		{16383, "XFD"},
	}
	for _, input := range table {
		got := indexToColumnName(input.i)
		if input.want != got {
			t.Errorf("Want %q, got %q", input.want, got)
		}
	}
}
