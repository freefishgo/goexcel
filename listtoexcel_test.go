package goexcel

import (
	"fmt"
	"testing"
	"time"
)

type s1 struct {
	Name string `json:"name" export:"一级姓名|姓名2,3"`
	Age  int32  `json:"age" export:"年龄2,6"`
	Time string `json:"time" export:"时间2,9"`
}

type s struct {
	Name string `json:"name" export:"一级姓名|姓名1,2" headStyle:"{\"fill\":{\"type\":\"pattern\",\"color\":[\"#E0EBF5\"],\"pattern\":1}}"`
	Age  int32  `json:"age" export:"年龄1,5" cellStyle:"{\"fill\":{\"type\":\"gradient\",\"color\":[\"#FFFFFF\",\"#E0EBF5\"],\"shading\":1}}"`
	Time string `json:"time" export:"时间1,8"`
	List []s1
}

type p struct {
	Name string `json:"name" export:"一级姓名|二级姓名|姓名,1"`
	Age  int32  `json:"age" export:"年龄,4"`
	Time string `json:"time" export:"时间,7" headStyle:"{\"font\":{\"bold\":true,\"italic\":true,\"family\":\"Berlin Sans FB Demi\",\"size\":36,\"color\":\"#777777\"}}"`
	List []s
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
	xlsx, err := ListToExcelSheet1(list)
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
	rows := func(row int) (style string) {
		if row%2 == 0 {
			return `{"fill":{"type":"pattern","color":["RED"],"pattern":1}}`
		}
		return ""
	}
	cell := func(cell int, value interface{}) (style string, newValue interface{}) {
		if cell == 9 {
			return `{"fill":{"type":"pattern","color":["RED"],"pattern":1}}`, value
		}
		return style, value
	}
	xlsx, err := ListToExcelSheet1Base(list, rows, cell)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	err = xlsx.SaveAs(time.Now().Format("20060102150405") + ".xlsx")
}
