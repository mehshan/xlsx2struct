package xlsx2struct

import (
	"fmt"
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/require"
	xlsx3 "github.com/tealeg/xlsx/v3"
)

type basicTypes struct {
	Bool    bool
	Float32 float32
	Float64 float64
	Int     int
	Int8    int8
	Int16   int16
	Int32   int32
	Int64   int64
	String  string
	Uint    uint
	Uint8   uint8
	Uint16  uint16
	Uint32  uint32
	Uint64  uint64
}

type supportedTypes struct {
	Bool    bool
	Float32 float32
	Float64 float64
	Int     int
	Int8    int8
	Int16   int16
	Int32   int32
	Int64   int64
	String  string
	Time    time.Time
	Uint    uint
	Uint8   uint8
	Uint16  uint16
	Uint32  uint32
	Uint64  uint64

	TrimString string `column:"trim"`

	Unsupported []int
}

func TestUnmarshalFieldBasicTypes(t *testing.T) {
	fields, err := fields(basicTypes{})
	require.NoError(t, err)

	for _, f := range fields {
		v, ok, err := unmarshalField(f, cell("1"))
		require.NoError(t, err)
		require.True(t, ok)
		require.NotNil(t, v)
		if f.Type == reflect.TypeOf(true) {
			require.Equal(t, true, v)
		} else {
			require.Equal(t, "1", fmt.Sprint(v))
		}
	}
}

func TestUnmarshalField(t *testing.T) {
	fields, err := fields(supportedTypes{})
	require.NoError(t, err)

	type test struct {
		Field    *Field
		Cell     *xlsx3.Cell
		Value    any
		ReadFlag bool
		Error    error
	}

	tests := []*test{
		{Field: fields["Unsupported"], Cell: cell(""), Error: &UnsupportedFieldError{Field: fields["Unsupported"]}},
		// no trim
		{Field: fields["String"], Cell: cell(" No Trim "), Value: " No Trim ", ReadFlag: true},
		// trim
		{Field: fields["TrimString"], Cell: cell(" Trim "), Value: "Trim", ReadFlag: true},
		// nil field and cell
		{Field: nil, Cell: nil, Error: &UnmarshalFieldError{Cell: nil, Field: nil}},
		// nil cell
		{Field: fields["String"], Cell: nil, Error: &UnmarshalFieldError{Cell: nil, Field: fields["String"]}},
	}

	for _, test := range tests {
		v, f, err := unmarshalField(test.Field, test.Cell)
		if test.Error == nil {
			require.NoError(t, err)
			require.Equal(t, test.Value, v)
			require.Equal(t, test.ReadFlag, f)
		} else {
			require.Error(t, err)
			require.EqualError(t, err, test.Error.Error())
			require.False(t, f)
		}
	}
}

func fields(a any) (map[string]*Field, error) {
	fs, err := extractFields(reflect.TypeOf(a))
	if err != nil {
		return nil, err
	}
	m := map[string]*Field{}
	for _, f := range fs {
		m[f.Name] = f
	}
	return m, nil
}

func cell(v string) *xlsx3.Cell {
	return &xlsx3.Cell{Value: v, Row: &xlsx3.Row{}}
}
