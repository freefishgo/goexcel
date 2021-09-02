# goexcel
excel

``` go
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
	Name   string `json:"name" export:"一级姓名|二级姓名|姓名,1"`
	Age    int32  `json:"age" export:"年龄,4"`
	Time   string `json:"time" export:"时间,7"`
	List   []s
}
func main() {
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
	v2:=&p{
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

```

<table width="659" border="0" cellpadding="0" cellspacing="0" style="width:659.00pt;border-collapse:collapse;table-layout:fixed;">
<tbody>
<tr height="12.40" style="height:12.40pt;">
    <td class="xl65" height="12.40" width="157.45" colspan="3" >一级姓名</td>
    <td class="xl65" width="59.15" rowspan="3" >年龄</td>
    <td class="xl65" width="50.40" rowspan="3" >年龄1</td>
    <td class="xl65" width="50.40" rowspan="3" >年龄2</td>
    <td class="xl65" width="81.65" rowspan="3" >时间</td>
    <td class="xl65" width="116.65" rowspan="3" >时间1</td>
    <td class="xl65" width="143.30" rowspan="3" >时间2</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" height="12.40" >二级姓名</td>
    <td class="xl65" rowspan="2">姓名1</td>
    <td class="xl65" rowspan="2" >姓名2</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" height="12.40" >姓名</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" height="24.80" rowspan="2" >天外飞仙</td>
    <td class="xl65" >大名</td>
    <td class="xl65" >大名</td>
    <td class="xl65" rowspan="2" >18</td>
    <td class="xl65" >19</td>
    <td class="xl65" >19</td>
    <td class="xl65" rowspan="2" >我是时间</td>
    <td class="xl65" >我是大名时间</td>
    <td class="xl65" >我是大名时间</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" >小名</td>
    <td class="xl65"></td>
    <td class="xl65" >20</td>
    <td class="xl65"></td>
    <td class="xl65" ></td>
    <td class="xl65"></td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" height="49.60" rowspan="4" >天外飞仙</td>
    <td class="xl65" rowspan="2" >小名</td>
    <td class="xl65" >大名</td>
    <td class="xl65" rowspan="4" >16</td>
    <td class="xl65" rowspan="2" >20</td>
    <td class="xl65" >19</td>
    <td class="xl65" rowspan="4" >我是开始时间</td>
    <td class="xl65" rowspan="2" >我是小名时间</td>
    <td class="xl65" >我是大名时间</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" >大名</td>
    <td class="xl65" >19</td>
    <td class="xl65" >我是大名时间</td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" >小名</td>
    <td class="xl65"></td>
    <td class="xl65" >21</td>
    <td class="xl65"></td>
    <td class="xl65" >我是小名名时间2</td>
    <td class="xl65"></td>
   </tr>
   <tr height="12.40" style="height:12.40pt;">
    <td class="xl65" >小名</td>
    <td class="xl65"></td>
    <td class="xl65" >21</td>
    <td class="xl65"></td>
    <td class="xl65" >我是小名名时间2</td>
    <td class="xl65"></td>
   </tr>
  </tbody></table>

``` go
// load from excel
var list []*p
goexcel.ExcelSheet1ToListFromPath("20210814182854.xlsx", &list)
```

