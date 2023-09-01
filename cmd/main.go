package main

// Using gooxml module

import (
	"gooxml/excel"
	"os"
)

func main() {
	doc := excel.NewDocument()

	f, err := os.Create("sample.xlsx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	if err = doc.Save(f); err != nil {
		panic(err)
	}
}
