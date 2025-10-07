// Package xlsx2struct builds on top of github.com/tealeg/xlsx to unmarshal XLSX sheets to Go structs.
package xlsx2struct

import (
	"reflect"
	"strings"

	xlsx3 "github.com/tealeg/xlsx/v3"
)

type SheetOptions struct {
	Row     int
	Col     int
	DataRow int
}

func DefaultSheetOptions() *SheetOptions {
	return &SheetOptions{
		Row:     0,
		Col:     0,
		DataRow: 1,
	}
}

func Unmarshal(sheet *xlsx3.Sheet, a any, opt *SheetOptions) error {
	// a should be pointer to slice

	v := reflect.ValueOf(a)
	if v.Kind() != reflect.Pointer || v.IsNil() {
		return &InvalidUnmarshalError{reflect.TypeOf(a)}
	}
	if v.Elem().Kind() != reflect.Slice {
		return &InvalidUnmarshalError{reflect.TypeOf(a)}
	}

	t := v.Elem().Type().Elem()
	items, err := unmarshalStructs(t, sheet, opt)
	if err != nil {
		return err
	}

	s := reflect.MakeSlice(v.Elem().Type(), 0, len(items))
	for _, i := range items {
		s = reflect.Append(s, reflect.ValueOf(i))
	}

	v.Elem().Set(s)

	return nil
}

func unmarshalStructs(t reflect.Type, sheet *xlsx3.Sheet, opt *SheetOptions) ([]any, error) {
	if sheet == nil {
		return nil, nil
	}

	if opt == nil {
		opt = DefaultSheetOptions()
	}

	fields, err := mapStructToSheet(t, sheet, opt.Row, opt.Col)
	if err != nil {
		return nil, err
	}

	row := opt.DataRow
	items := []any{}

	for {
		values, ok, err := unmarshalFields(fields, sheet, row)
		if err != nil {
			return nil, err
		}

		if !ok {
			break // empty row found
		}

		item, err := newStruct(t, values)
		if err != nil {
			return nil, err
		}

		items = append(items, item)
		row += 1
	}

	return items, nil
}

// unmarshalStruct unmarshals fields from the given sheet row.
func unmarshalFields(fields map[*Field]*Column, sheet *xlsx3.Sheet, row int) (map[*Field]any, bool, error) {
	if sheet == nil || row < 0 || len(fields) == 0 {
		return nil, false, nil
	}

	m := map[*Field]any{}
	allOk := false

	for f, col := range fields {
		var c *xlsx3.Cell

		if col != nil {
			c, _ = sheet.Cell(row, col.Index) // TODO: ok to ignore error...?
		}

		v, ok, err := unmarshalField(f, c)
		if err != nil {
			return nil, false, err
		}

		m[f] = v

		if ok {
			allOk = true
		}
	}

	return m, allOk, nil
}

func mapStructToSheet(t reflect.Type, sheet *xlsx3.Sheet, row, col int) (map[*Field]*Column, error) {
	cols, err := extractColumns(sheet, row, col)
	if err != nil {
		return nil, err
	}

	fields, err := extractFields(t)
	if err != nil {
		return nil, err
	}

	return mapFields(fields, cols), nil
}

// mapFields maps fields of struct to sheet columns.
func mapFields(fields []*Field, columns []*Column) map[*Field]*Column {
	cs := map[string]*Column{}
	for _, c := range columns {
		cs[c.Heading] = c
	}

	fs := map[*Field]*Column{}
	for _, f := range fields {
		h := f.Heading()
		fs[f] = cs[h]
	}

	return fs
}

// t must be a struct type or pointer to a struct type.
func extractFields(t reflect.Type) ([]*Field, error) {
	var s reflect.Type

	if t.Kind() == reflect.Struct {
		s = t
	} else if t.Kind() == reflect.Pointer && t.Elem().Kind() == reflect.Struct {
		s = t.Elem()
	} else {
		return nil, &InvalidUnmarshalError{Type: t}
	}

	fs := make([]*Field, 0)

	for i := 0; i < s.NumField(); i++ {
		f := Field{StructField: s.Field(i)}
		if c := f.Tag.Get(ColumnTag); c != "" {
			f.tag = parseColumnTag(c)
		}
		fs = append(fs, &f)
	}

	return fs, nil
}

// field is mapped to a column
type Column struct {
	Heading string
	Index   int
}

func extractColumns(sheet *xlsx3.Sheet, row, col int) ([]*Column, error) {
	if sheet == nil || row < 0 || col < 0 {
		return nil, nil
	}

	cols := []*Column{}

	for {
		c, err := sheet.Cell(row, col)
		if err != nil {
			return nil, err
		}

		v := strings.TrimSpace(c.Value)
		if v == "" {
			break
		}

		cols = append(cols, &Column{Heading: c.Value, Index: col})
		col += 1
	}

	return cols, nil
}
