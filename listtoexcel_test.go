package goexcel_test

import (
	"fmt"
	"github.com/freefishgo/goexcel"
	"testing"
	"time"
)

type s1 struct {
	Name string `json:"name" export:"一级姓名|姓名2,3"`
	Age  int32  `json:"age" export:"年龄2,6"`
	Time string `json:"time" export:"时间2,9"`
}

type s struct {
	Name string `json:"name" export:"一级姓名|姓名1,2"`
	Age  int32  `json:"age" export:"年龄1,5"`
	Time string `json:"time" export:"时间1,8"`
	List []s1
}

type p struct {
	Name string `json:"name" export:"一级姓名|二级姓名|姓名,1"`
	Age  int32  `json:"age" export:"年龄,4"`
	Time string `json:"time" export:"时间,7"`
	List []s
}

func TestAxisToCellRow(t *testing.T) {
	fmt.Println(goexcel.AxisToCellRow("A12"))
}

func TestListToExcel(t *testing.T) {
	v := &p{
		Name: "天外飞仙",
		Age:  18,
		Time: "我是时间",
		List: []s{
			{
				Name: "大名",
				Age:  19,
				Time: "我是大名时间",
				List: []s1{
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
				},
			},
			{
				Name: "小名",
				Age:  20,
			},
		},
	}
	v2 := &p{
		Name: "天外飞仙",
		Age:  16,
		Time: "我是开始时间",
		List: []s{
			{
				Name: "小名",
				Age:  20,
				Time: "我是小名时间",
				List: []s1{
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
				},
			},
			{
				Name: "小名",
				Age:  21,
				Time: "我是小名名时间2",
			},
			{
				Name: "小名",
				Age:  21,
				Time: "我是小名名时间2",
			},
		},
	}
	list := append([]*p(nil), v, v2)
	xlsx, err := goexcel.ListToExcelSheet1(list)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	err = xlsx.SaveAs(time.Now().Format("20060102150405") + ".xlsx")
}

func TestListToExcelSheet1Base(t *testing.T) {
	v := &p{
		Name: "天外飞仙",
		Age:  18,
		Time: "我是时间",
		List: []s{
			{
				Name: "大名",
				Age:  19,
				Time: "我是大名时间",
				List: []s1{
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
				},
			},
			{
				Name: "小名",
				Age:  20,
			},
		},
	}
	list := append([]*p(nil), &p{
		Name: "天外飞仙",
		Age:  16,
		Time: "我是开始时间",
		List: []s{
			{
				Name: "小名",
				Age:  20,
				Time: "我是小名时间",
				List: []s1{
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
					{
						Name: "大名",
						Age:  19,
						Time: "我是大名时间",
					},
				},
			},
			{
				Name: "小名",
				Age:  21,
				Time: "我是小名名时间2",
			},
			{
				Name: "小名",
				Age:  21,
				Time: "我是小名名时间2",
			},
		},
	}, v, v, v, v, v, v, v, v, v, v)
	//rows := func(row int) (style string) {
	//	if row%2 == 0 {
	//		return `{"fill":{"type":"pattern","color":["RED"],"pattern":1}}`
	//	}
	//	return ""
	//}
	//cell := func(cell int, value interface{}) (style string, newValue interface{}) {
	//	if cell == 9 {
	//		return `{"fill":{"type":"pattern","color":["RED"],"pattern":1}}`, value
	//	}
	//	return style, value
	//}
	xlsx, err := goexcel.ListToExcelSheet1Base(list, nil)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	err = xlsx.SaveAs(time.Now().Format("tmp/"+"20060102150405") + ".xlsx")
}
