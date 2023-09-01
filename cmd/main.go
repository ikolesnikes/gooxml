package main

// Using gooxml module

import (
	"os"

	"github.com/ikolesnikes/gooxml/excel"
)

func main() {
	doc := excel.NewDocument()

	f, err := os.Create("/tmp/sample.xlsx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	if err = doc.Save(f); err != nil {
		panic(err)
	}
}
