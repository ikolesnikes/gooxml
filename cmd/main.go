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
	wks.AddText("n", 0, 0)
	wks.AddText("n\xc2\xb2", 0, 1)
	wks.AddText("n\xc2\xb3", 0, 2)
	for i := 1; i <= 30; i++ {
		wks.AddText(fmt.Sprintf("%d", i), i, 0)
		wks.AddText(fmt.Sprintf("%d", i*i), i, 1)
		wks.AddText(fmt.Sprintf("%d", i*i*i), i, 2)
	}

	f, err := os.Create("/tmp/sample.xlsx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	if err = doc.Save(f); err != nil {
		panic(err)
	}
}
