package xlsx2struct

import "reflect"

func newStruct(t reflect.Type, values map[*Field]any) (any, error) {
	s, ptr := getStructType(t)
	if s == nil {
		return nil, &InvalidUnmarshalError{Type: t}
	}

	v := reflect.New(s).Elem()

	for field, value := range values {
		f := v.FieldByName(field.Name)
		if !f.CanSet() {
			return nil, &InvalidFieldError{Field: field}
		}

		fv := reflect.ValueOf(value)
		if f.Type() != fv.Type() {
			return nil, &InvalidFieldValueError{Field: field, Value: value}
		}

		f.Set(fv)
	}

	var a any

	if ptr {
		pt := reflect.PointerTo(s)
		pv := reflect.New(pt.Elem())
		pv.Elem().Set(v)
		a = pv.Interface()
	} else {
		a = v.Interface()
	}

	return a, nil
}

func getStructType(t reflect.Type) (reflect.Type, bool) {
	var s reflect.Type
	ptr := false

	if t.Kind() == reflect.Struct {
		s = t
	} else if t.Kind() == reflect.Pointer && t.Elem().Kind() == reflect.Struct {
		s = t.Elem()
		ptr = true
	} else {
		return nil, false
	}

	return s, ptr
}
