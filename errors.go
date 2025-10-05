package xlsx2struct

import (
	"fmt"
	"reflect"
	"strconv"

	xlsx3 "github.com/tealeg/xlsx/v3"
)

type InvalidUnmarshalError struct {
	Type reflect.Type
}

func (e *InvalidUnmarshalError) Error() string {
	if e.Type == nil {
		return "xlsx2struct: invalid unmarshal(nil)"
	}
	if e.Type.Kind() != reflect.Pointer {
		return "xlsx2struct: invalid unmarshal(non-pointer " + e.Type.String() + ")"
	}
	return "xlsx2struct: invalid unmarshal(nil " + e.Type.String() + ")"
}

type UnmarshalFieldError struct {
	Field *Field
	Cell  *xlsx3.Cell
}

func (e *UnmarshalFieldError) Error() string {
	return "xlsx2struct: cannot unmarshal cell " + describeCell(e.Cell) + " into field " + e.Field.Describe()
}

type UnsupportedFieldError struct {
	Field  *Field
	Column *Column
}

func (e *UnsupportedFieldError) Error() string {
	f := "nil"
	if e.Field != nil {
		h := e.Field.Name
		if e.Column != nil {
			h = e.Column.Heading
		}
		f = fmt.Sprintf("'%s' (%s)", h, e.Field.Type)
	}
	return "xlsx2struct: unsupported field " + strconv.Quote(f)
}

type UnsupportedValueError struct {
	Value string
}

func (e *UnsupportedValueError) Error() string {
	return "xlsx2struct: unsupported value " + strconv.Quote(e.Value)
}

type InvalidFieldError struct {
	Field *Field
}

func (e *InvalidFieldError) Error() string {
	return "xlsx2struct: invalid field " + e.Field.Describe()
}

type InvalidFieldValueError struct {
	Field *Field
	Value any
}

func (e *InvalidFieldValueError) Error() string {
	return "xlsx2struct: invalid value " + describe(e.Value) + " for field " + e.Field.Describe()
}

func describe(a any) string {
	return fmt.Sprintf("'%v' (type: %v)", a, reflect.TypeOf(a))
}

func describeCell(c *xlsx3.Cell) string {
	s := "nil"
	if c != nil {
		x, y := c.GetCoordinates()
		s = fmt.Sprintf("(%d, %d)[%s]", x, y, c.String())
	}
	return s
}
