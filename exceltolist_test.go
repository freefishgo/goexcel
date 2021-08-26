package goexcel_test

import (
	"github.com/freefishgo/goexcel"
	"reflect"
	"testing"
)

func TestNewExcelSheet1ToListFromPath(t *testing.T) {
	goexcel.NewExcelSheet1ToListFromPath("20210814182854.xlsx", reflect.TypeOf(&p{}))
}
