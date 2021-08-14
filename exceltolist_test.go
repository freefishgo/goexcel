package goexcel

import (
	"reflect"
	"testing"
)

func TestNewExcelSheet1ToListFromPath(t *testing.T) {
	NewExcelSheet1ToListFromPath("20210814182854.xlsx", reflect.TypeOf(&p{}))
}
