package main

// Using gooxml module

import (
	"fmt"
	"os"

	"github.com/ikolesnikes/gooxml/excel"
)

func main() {
	doc := excel.NewDocument()

	// The newly created document already contains a workbook
	// and a worksheet.

	wks := doc.Workbook().Worksheet(0)

	wks.SetText("n", 0, 0)
	wks.SetText("n\xc2\xb2", 0, 1)
	wks.SetText("n\xc2\xb3", 0, 2)
	for i := 1; i <= 30; i++ {
		x := i
		wks.SetText(fmt.Sprintf("%d", x), i, 0)
		x *= i
		wks.SetText(fmt.Sprintf("%d", x), i, 1)
		x *= i
		wks.SetText(fmt.Sprintf("%d", x), i, 2)
	}

	f, err := os.Create("sample.xlsx")
	if err != nil {
		panic(err)
	}

	if err = doc.Save(f); err != nil {
		panic(err)
	}
	f.Close()

	/*wks.SetText("After save", 0, 4)

	f, err = os.Create("sample.xlsx")
	if err != nil {
		panic(err)
	}

	if err = doc.Save(f); err != nil {
		panic(err)
	}
	f.Close()*/
}
