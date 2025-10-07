# xlsx2struct

Go XLSX to struct library.

## Import the package

To import the package, use the line

`import "github.com/mehshan/xlsx2struct"`

## Example

This example uses sample [salesorders.xlsx](testdata/salesorders.xlsx) spreadsheet.

```
import (
    "time"

    github.com/mehshan/xlsx2struct
)

type SaleOrder struct {
    Date   time.Time `column:"heading=Order Date,trim,time=2004-01-01,default=o"`
}

type SaleOrder struct {
    Date   time.Time `heading:"Order Date"`
    Region string    `heading:"Region"`
    Rep    string    `heading:"Rep"`
    Item   string    `heading:"Item"`
    Units  int32     `heading:"Units"`
    Cost   float32   `heading:"Unit Cost"`
    Total  float64   `heading:"Total"`
}

func main() {
    opt := DefaultSheetOptions()

    orders := []*SaleOrder{}
    err := Unmarshal(sheet, &orders, opt)

    if err != nil {
        panic(err)
    }

    for _, o := range orders {
        fmt.Println("%v %s %s %s %d %f %f", o.Date, o.Region, o.Rep, o.Item, o.Units, o.Cost, o.Total)
    }
}

```