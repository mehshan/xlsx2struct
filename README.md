# xlsx2struct

[[https://github.com/mehshan/xlsx2struct/blob/main/LICENSE][https://img.shields.io/badge/license-bsd-orange.svg]]

**xlsx2struct** builds on top of [github.com/tealeg/xlsx](https://github.com/tealeg/xlsx) to unmarshal XLSX sheets to Go structs.

## Import the package

To import the package, use the line

`import "github.com/mehshan/xlsx2struct"`

## Example

This example uses sample [salesorders.xlsx](testdata/salesorders.xlsx) spreadsheet. Complete code is available in [examples repo](https://github.com/mehshan/xlsx2struct-examples/tree/main/salesorders).

```go
type SaleOrder struct {
	Date   time.Time `column:"heading=Order Date"`
	Region string    `column:"heading=Region,trim"`
	Rep    string    `column:"heading=Rep"`
	Item   string    `column:"heading=Item,default=Pencil"`
	Units  int32     `column:"heading=Units,default=1"`
	Cost   float32   `column:"heading=Unit Cost"`
	Total  float64   `column:"heading=Total"`
}

func main() {
	orders := []*SaleOrder{}

	file, err := xlsx3.OpenFile("testdata/salesorders.xlsx")
	if err != nil {
		panic(err)
	}

	sheet := file.Sheet["Sales Orders"]
	opt := xlsx2struct.DefaultSheetOptions()

	err = xlsx2struct.Unmarshal(sheet, &orders, opt)
	if err != nil {
		panic(err)
	}

	for _, o := range orders {
		fmt.Printf("%v %s %s %s %d %f %f\n", o.Date, o.Region, o.Rep, o.Item, o.Units, o.Cost, o.Total)
	}
}

```