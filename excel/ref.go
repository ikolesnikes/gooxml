package excel

import "fmt"

// makeRef creates the A1 style reference out of given row and
// column zero-based indexes.
func makeRef(ri, ci int) string {
	return fmt.Sprintf("%s%d", indexToColumnName(ci), ri+1)
}

// indexToColumnName converts a zero-based index into Excel's
// column name (A, B, etc.).
//
// Maximum column in the .xlsx file is XFD (16383).
func indexToColumnName(i int) string {
	var name string
	var r int
	for {
		r, i = i%26, i/26-1
		name = fmt.Sprintf("%c%s", r+'A', name)
		if i < 0 {
			break
		}
	}
	return name
}
