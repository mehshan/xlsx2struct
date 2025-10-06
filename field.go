package xlsx2struct

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	xlsx3 "github.com/tealeg/xlsx/v3"
)

var defaultTimeFormats = []string{time.DateOnly, time.RFC3339}

type Field struct {
	reflect.StructField
	tag columnTag
}

func (f Field) Heading() string {
	if f.tag.heading != "" {
		return f.tag.heading
	}
	return f.Name
}

func (f Field) String() string {
	return f.Heading()
}

func (f *Field) Describe() string {
	s := "nil"
	if f != nil {
		h := f.Name
		if f.tag.heading != "" {
			h = f.tag.heading
		}
		t := "nil"
		if f.Type != nil {
			t = f.Type.String()
		}
		s = fmt.Sprintf("'%s' (type: %s, column: '%s')", f.Name, t, h)
	}
	return s
}

func unmarshalField(field *Field, cell *xlsx3.Cell) (a any, ok bool, err error) {
	if field == nil || cell == nil {
		err = &UnmarshalFieldError{Cell: cell, Field: field}
		return
	}

	v := cell.Value
	ok = true

	if v == "" {
		v = defaultValue(field)
		ok = false
	}

	if field.tag.trim {
		v = strings.TrimSpace(v)
	}

	var i int64
	var u uint64
	var f float64
	var t time.Time

	switch field.Type.Kind() {
	case reflect.Bool:
		a, err = strconv.ParseBool(v)
	case reflect.Float32:
		if f, err = strconv.ParseFloat(v, 32); err == nil {
			a = float32(f)
		}
	case reflect.Float64:
		if f, err = strconv.ParseFloat(v, 64); err == nil {
			a = f
		}
	case reflect.Int:
		if i, err = strconv.ParseInt(v, 10, 64); err == nil {
			a = int(i)
		}
	case reflect.Int8:
		if i, err = strconv.ParseInt(v, 10, 8); err == nil {
			a = int8(i)
		}
	case reflect.Int16:
		if i, err = strconv.ParseInt(v, 10, 16); err == nil {
			a = int16(i)
		}
	case reflect.Int32:
		if i, err = strconv.ParseInt(v, 10, 32); err == nil {
			a = int32(i)
		}
	case reflect.Int64:
		a, err = strconv.ParseInt(v, 10, 64)
	case reflect.Uint:
		if u, err = strconv.ParseUint(v, 10, 64); err == nil {
			a = uint(u)
		}
	case reflect.Uint8:
		if u, err = strconv.ParseUint(v, 10, 8); err == nil {
			a = uint8(u)
		}
	case reflect.Uint16:
		if u, err = strconv.ParseUint(v, 10, 16); err == nil {
			a = uint16(u)
		}
	case reflect.Uint32:
		if u, err = strconv.ParseUint(v, 10, 32); err == nil {
			a = uint32(u)
		}
	case reflect.Uint64:
		a, err = strconv.ParseUint(v, 10, 64)
	case reflect.String:
		a = v
	case reflect.Struct:
		switch field.Type {
		case reflect.TypeOf(time.Time{}): // TODO: re-factor unmarshalTimeField
			switch cell.Type() {
			case xlsx3.CellTypeNumeric:
				if f, err = strconv.ParseFloat(v, 64); err == nil {
					a = xlsx3.TimeFromExcelTime(f, false)
				}
			default:
				if t, err = parseTime(v, field.tag.timeFormats...); err == nil {
					a = t
				}
			}
		default:
			err = &UnsupportedFieldError{Field: field}
		}
	default:
		err = &UnsupportedFieldError{Field: field}
	}

	if err != nil {
		err = &UnmarshalFieldError{Cell: cell, Field: field}
	}

	return
}

func parseTime(v string, formats ...string) (time.Time, error) {
	if len(formats) == 0 {
		formats = defaultTimeFormats
	}

	for _, f := range formats {
		if t, err := time.Parse(f, v); err == nil {
			return t, nil
		}
	}

	return time.Time{}, &UnsupportedValueError{Value: v}
}

func defaultValue(f *Field) string {
	if f == nil {
		return ""
	}

	v := f.tag.defaultValue
	if v == "" {
		switch f.Type.Kind() {
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64, reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			v = "0"
		case reflect.Float32, reflect.Float64:
			v = "0.0"
		case reflect.Struct:
			switch f.Type {
			case reflect.TypeOf(time.Time{}):
				v = "0001-01-01"
			}
		}
	}

	return v
}
