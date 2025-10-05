package xlsx2struct

import (
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/require"
	xlsx "github.com/tealeg/xlsx/v3"
)

type SaleOrder struct {
	Date   time.Time `column:"heading=Order Date"`
	Region string    `column:"heading=Region"`
	Rep    string    `column:"heading=Rep"`
	Item   string    `column:"heading=Item,default=Pencil"`
	Units  int32     `column:"heading=Units,default=1"`
	Cost   float32   `column:"heading=Unit Cost"`
	Total  float64   `column:"heading=Total"`
}

func TestUnmarshal(t *testing.T) {
	sheet, _ := openSalesOrdersSheet(t)

	a := []*SaleOrder{}
	a = append(a, &SaleOrder{})

	opt := DefaultSheetOptions()
	opt.DataRow = 1

	err := Unmarshal(sheet, &a, opt)
	require.NoError(t, err)
	require.Len(t, a, 20)

	// dates
	require.Equal(t, "2021-01-06", a[0].Date.Format(time.DateOnly))

	// default values
	require.Equal(t, int32(1), a[1].Units)
	require.Equal(t, "Pencil", a[2].Item)
}

func TestUnmarshalFields(t *testing.T) {
	sheet, _ := openSalesOrdersSheet(t)

	type Struct1 = SaleOrder
	fields, err := mapStructToSheet(reflect.TypeOf(Struct1{}), sheet, 0, 0)
	require.NoError(t, err)

	// empty row
	_, ok, err := unmarshalFields(fields, sheet, 25)
	require.NoError(t, err)
	require.False(t, ok)

	// data row
	a, ok, err := unmarshalFields(fields, sheet, 1)
	require.NoError(t, err)
	require.True(t, ok)
	require.NotNil(t, a)
	require.Equal(t, len(fields), len(a))
}

func TestMapFields(t *testing.T) {
	// TODO: complete test case
}

func TestExtractColumns(t *testing.T) {
	sheet, _ := openSalesOrdersSheet(t)
	cols, err := extractColumns(sheet, 0, 0)
	require.NoError(t, err)
	require.Equal(t, 7, len(cols))
	require.Equal(t, "Order Date", cols[0].Heading)
	require.Equal(t, 0, cols[0].Index)
	require.Equal(t, "Total", cols[6].Heading)
	require.Equal(t, 6, cols[6].Index)
}

func openSalesOrdersSheet(t *testing.T) (*xlsx.Sheet, *SheetOptions) {
	sheet := openSheet(t, "testdata/salesorders.xlsx", "Sales Orders")
	opt := DefaultSheetOptions()
	opt.DataRow = 1
	return sheet, opt
}

func openSheet(t *testing.T, path, sheet string) *xlsx.Sheet {
	f, err := xlsx.OpenFile(path)
	require.NoError(t, err)
	s := f.Sheet[sheet]
	require.NotNil(t, s)
	return s
}
