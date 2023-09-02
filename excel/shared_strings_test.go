package excel

import "testing"

func Test_indexDistinctItemsAdded(t *testing.T) {
	sst := newSharedStrings()
	table := []struct {
		text string
		want int
	}{
		{"foo1", 0},
		{"foo2", 1},
		{"foo3", 2},
	}
	for _, input := range table {
		got := sst.add(input.text)
		if input.want != got {
			t.Errorf("Want: %q, got %q", input.want, got)
		}
	}
}

func Test_indexSameItemsAdded(t *testing.T) {
	sst := newSharedStrings()
	table := []struct {
		text string
		want int
	}{
		{"foo", 0},
		{"foo", 0},
		{"foo", 0},
	}
	for _, input := range table {
		got := sst.add(input.text)
		if input.want != got {
			t.Errorf("Want: %q, got %q", input.want, got)
		}
	}
}
