package excel

import "testing"

func Test_StringItemIndex(t *testing.T) {
	sst := newSharedStrings()
	table := []struct {
		text string
		want int
	}{
		{"foo", 0},
		{"foo2", 1},
		{"foo3", 2},
		{"foo", 0},
		{"foo4", 3},
		{"foo2", 1},
	}
	for _, input := range table {
		got := sst.add(input.text)
		if input.want != got {
			t.Errorf("Want: %q, got %q", input.want, got)
		}
	}
}

func Test_StringItemCount(t *testing.T) {
	sst := newSharedStrings()
	table := []struct {
		text string
		want int
	}{
		{"foo", 3},
		{"foo", 3},
		{"foo2", 2},
		{"foo3", 1},
		{"foo", 3},
		{"foo4", 1},
		{"foo2", 2},
	}
	for _, input := range table {
		_ = sst.add(input.text)
	}
	for _, input := range table {
		got := sst.items[input.text].count
		if input.want != got {
			t.Errorf("Want: %q, got %q", input.want, got)
		}
	}
}
