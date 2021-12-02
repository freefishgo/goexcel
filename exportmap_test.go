package goexcel_test

import (
	"fmt"
	"github.com/freefishgo/goexcel"
	"testing"
	"time"
)

type m1 struct {
	Name string `json:"name" export:"一级姓名|二级姓名|姓名,1"`
	Age  int32  `json:"age" export:"年龄,4"`
	Time string `json:"time" export:"时间,7"`
	Map  *goexcel.ExportMap[string, *s]
}

func TestNewExportMap(t *testing.T) {
	e := goexcel.NewExportMap[string, *m1]()
	v := &m1{
		Name: "test",
		Map:  goexcel.NewExportMap[string, *s](),
	}
	v.Map.SetValue("map1", &s{
		Age:  998,
		Name: "s名",
	})
	v.Map.SetValue("map2", &s{
		Age:  9989,
		Name: "s名9",
	})
	v.Map.SetValue("map3", &s{
		Age:  9989,
		Name: "s名9",
	})
	e.SetValue("KEY1", v)
	e.SetValue("KEY2", v)
	e.SetValue("KEY3", v)
	xlsx, err := goexcel.ListToExcelSheet1Base(e, nil)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	err = xlsx.SaveAs(time.Now().Format("tmp/"+"20060102150405") + ".xlsx")

}
