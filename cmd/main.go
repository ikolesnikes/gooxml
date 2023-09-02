package main

// Using gooxml module

import (
	"os"

	"github.com/ikolesnikes/gooxml/excel"
)

func main() {
	doc := excel.NewDocument()

	// The newly created document already contains a workbook
	// and a worksheet.

	wks := doc.Workbook().Worksheet(0)
	wks.AddText("foo", 0, 0)
	wks.AddText("bar", 0, 1)

	// wkb := doc.Workbook()
	// Add second worksheet
	// wkb.AddWorksheet()

	f, err := os.Create("/tmp/sample.xlsx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	if err = doc.Save(f); err != nil {
		panic(err)
	}
}
