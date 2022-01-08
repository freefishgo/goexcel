package goexcel

import (
	"encoding/json"
	"reflect"
	"strconv"
)

func stringSetValue(des reflect.Value, val string) {
	if val == "" {
		return
	}
	des = getBaseValue(des)
	switch des.Type().Kind() {
	case reflect.String:
		if des.CanSet() {
			des.SetString(val)
		}
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		if val, err := strconv.ParseInt(val, 10, des.Type().Bits()); err == nil {
			if des.CanSet() {
				des.SetInt(val)
			}
		}
	case reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		if val, err := strconv.ParseUint(val, 10, des.Type().Bits()); err == nil {
			if des.CanSet() {
				des.SetUint(val)
			}
		}
	case reflect.Bool:
		if val, err := strconv.ParseBool(val); err == nil {
			if des.CanSet() {
				des.SetBool(val)
			}
		}
	case reflect.Float32, reflect.Float64:
		if val, err := strconv.ParseFloat(val, des.Type().Bits()); err == nil {
			if des.CanSet() {
				des.SetFloat(val)
			}
		}
	case reflect.Struct, reflect.Slice:
		vTemp := reflect.New(des.Type()).Interface()
		err := json.Unmarshal([]byte(val), vTemp)
		if err == nil {
			des.Set(reflect.ValueOf(vTemp).Elem())
		}
	case reflect.Interface:
		if des.CanSet() {
			des.Set(reflect.ValueOf(val))
		}
	}
}
