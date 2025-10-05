package xlsx2struct

import (
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/require"
)

type basics struct {
	Bool   bool
	Int    int
	String string
	Uint   uint
}

type ints struct {
	Int8  int8
	Int16 int16
	Int32 int32
	Int64 int64
}

type uints struct {
	Uint8  uint8
	Uint16 uint16
	Uint32 uint32
	Uint64 uint64
}

type floats struct {
	Float32 float32
	Float64 float64
}

type wellknown struct {
	Time time.Time
}

type unexported struct {
	int    int
	string string
}

type mixed struct {
	Int    int
	String string
	int    int
	string string
}

func TestNewStruct(t *testing.T) {
	type test struct {
		Any         any
		FieldValues map[*Field]any
		Error       string
	}

	tests := []*test{
		{
			Any:         basics{Bool: true, Int: -110, String: "abc", Uint: 210},
			FieldValues: mapFieldValuePairs("Bool", true, "Int", int(-110), "String", "abc", "Uint", uint(210)),
		},
		{
			Any:         &basics{Bool: false, Int: 1000, String: "xyz", Uint: 7},
			FieldValues: mapFieldValuePairs("Bool", false, "Int", int(1000), "String", "xyz", "Uint", uint(7)),
		},
		{
			Any:         ints{Int8: -100, Int16: 3200, Int32: -43203, Int64: 32483432},
			FieldValues: mapFieldValuePairs("Int8", int8(-100), "Int16", int16(3200), "Int32", int32(-43203), "Int64", int64(32483432)),
		},
		{
			Any:         uints{Uint8: 255, Uint16: 13200, Uint32: 643203, Uint64: 532483432},
			FieldValues: mapFieldValuePairs("Uint8", uint8(255), "Uint16", uint16(13200), "Uint32", uint32(643203), "Uint64", uint64(532483432)),
		},
		{
			Any:         floats{Float32: 44.32, Float64: 2203.8789},
			FieldValues: mapFieldValuePairs("Float32", float32(44.32), "Float64", float64(2203.8789)),
		},
		{
			Any:         wellknown{Time: time.UnixMicro(1733063813)},
			FieldValues: mapFieldValuePairs("Time", time.UnixMicro(1733063813)),
		},
		{
			Any:         mixed{String: "Exported", Int: 10, string: "abc", int: 20},
			FieldValues: mapFieldValuePairs("String", "Exported", "Int", 10),
		},
		{
			Error:       "xlsx2struct: invalid field 'unexported' (type: int, column: 'unexported')",
			Any:         unexported{int: 10, string: "abc"},
			FieldValues: mapFieldValuePairs("unexported", 10),
		},
		{
			Error:       "xlsx2struct: invalid value '101' (type: int) for field 'String' (type: int, column: 'String')",
			Any:         basics{Bool: true, Int: -110, String: "abc", Uint: 210},
			FieldValues: mapFieldValuePairs("String", 101),
		},
		{
			Error:       "xlsx2struct: invalid unmarshal(non-pointer string)",
			Any:         "string",
			FieldValues: mapFieldValuePairs("String", 101),
		},
	}

	for _, test := range tests {
		a, err := newStruct(reflect.TypeOf(test.Any), test.FieldValues)
		if err == nil && test.Error == "" {
			require.NoError(t, err)
			require.NotNil(t, a)
			require.EqualExportedValues(t, test.Any, a)
		} else {
			require.Error(t, err)
			require.EqualError(t, err, test.Error)
		}
	}
}

func mapFieldValuePairs(a ...any) map[*Field]any {
	m := map[*Field]any{}
	for i := 0; i < len(a); i += 2 {
		n := a[i].(string)
		v := a[i+1]
		f := &Field{StructField: reflect.StructField{Name: n, Type: reflect.TypeOf(v)}}
		m[f] = v
	}
	return m
}
