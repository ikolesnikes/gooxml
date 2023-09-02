# gooxml

## Example

```go
doc := excel.NewDocument()

// The newly created document already contains a workbook
// and a worksheet.

wks := doc.Workbook().Worksheet(0)
wks.AddText("foo", 0, 0)

// Add second worksheet
doc.Workbook().AddWorksheet()
wks = doc.Workbook().Worksheet(1)
wks.AddText("bar", 3, 4)

f, err := os.Create("/tmp/sample.xlsx")
if err != nil {
    panic(err)
}
defer f.Close()

if err = doc.Save(f); err != nil {
    panic(err)
}
```
