package goexcel

import (
	"encoding/json"
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
)

// GetCellCode 获取excel列标号
func GetCellCode(columnNumber int) string {
	var ans []byte
	for columnNumber > 0 {
		columnNumber--
		ans = append(ans, 'A'+byte(columnNumber%26))
		columnNumber /= 26
	}
	for i, n := 0, len(ans); i < n/2; i++ {
		ans[i], ans[n-1-i] = ans[n-1-i], ans[i]
	}
	return string(ans)
}

// AxisToCellRow Axis转换成列号和行号
func AxisToCellRow(axis string) (cell string, row int) {
	end := 1
	for ; end < len(axis); end++ {
		if axis[end] >= '0' || axis[end] <= '9' {
			break
		}
	}
	cell = axis[:end]
	row, _ = strconv.Atoi(axis[end:])
	return
}

// ListToExcelSheet1 通过list数据来生成数据格式 通过结构体 tag export:"一级姓名|姓名1,2"
//
// 生成表头和数据格式 通过,来分割表头名和列位置  表头名通过 |来判断层级
//
// 数据级转换成 excelize.File
func ListToExcelSheet1(list interface{}) (*excelize.File, error) {
	return ListToExcelSheet1Base(list, nil, nil)
}

// ListToExcelSheet1Base 通过list数据来生成数据格式 通过结构体 tag export:"一级姓名|姓名1,2"
//
// 生成表头和数据格式 通过,来分割表头名和列位置  表头名通过 |来判断层级
//
// rowStyle 对list行处理 cellDo 单元格处理
func ListToExcelSheet1Base(list interface{}, rowStyle func(row int) (style string), cellDo func(cell int, value interface{}) (style string, newValue interface{})) (*excelize.File, error) {
	base, err := getSliceBaseType(list)
	if err != nil {
		return nil, err
	}
	export := getExportSort(base, true)
	if len(export.allFields) == 0 {
		return nil, errors.New("结构体没有导出的字段")
	}
	arr := reflect.ValueOf(list)
	xlsx := excelize.NewFile()
	tableName := "Sheet1"
	hashStyle := map[string]int{}
	cellMap := make([]string, len(export.allFields)+2)
	cellStyleList := make([]int, len(export.allFields)+1)
	haveStyle := false
	for _, v := range export.allFields {
		cell := GetCellCode(v.sort)
		cellMap[v.sort] = cell
		for row, v1 := range v.headInfo {
			xlsx.SetCellValue(tableName, cell+strconv.Itoa(row+1), v1)
		}
		if len(v.headInfo) < export.headMaxLevel {
			xlsx.MergeCell(tableName, cell+strconv.Itoa(len(v.headInfo)), cell+strconv.Itoa(export.headMaxLevel))
		}
		if v.headStyle != "" {
			haveStyle = true
		}
		if v.cellStyle != "" {
			haveStyle = true
			cellStyleList[v.sort], err = xlsx.NewStyle(v.cellStyle)
			hashStyle[v.cellStyle] = cellStyleList[v.sort]
			if err != nil {
				return nil, err
			}
		}
	}
	cellMap[len(export.allFields)+1] = GetCellCode(len(export.allFields) + 1)
	for i := 1; i <= export.headMaxLevel; i++ {
		startJ := 1
		rowS := strconv.Itoa(i)
		startJv := xlsx.GetCellValue(tableName, cellMap[1]+rowS)
		for j := 2; j <= len(export.allFields)+1; j++ {
			endJ := xlsx.GetCellValue(tableName, cellMap[j]+rowS)
			if startJv != endJ {
				mgEnd := i
				mg := export.allFields[j-2]
				style := 0
				if len(mg.headInfo) == i {
					mgEnd = export.headMaxLevel
					if mg.headStyle != "" {
						ok := false
						style, ok = hashStyle[mg.headStyle]
						if !ok {
							style, err = xlsx.NewStyle(mg.headStyle)
							if err != nil {
								return xlsx, err
							}
							hashStyle[mg.headStyle] = style
						}
						xlsx.SetCellStyle(tableName, cellMap[startJ]+rowS, cellMap[j-1]+strconv.Itoa(mgEnd), style)
					}
				}
				xlsx.MergeCell(tableName, cellMap[startJ]+rowS, cellMap[j-1]+strconv.Itoa(mgEnd))
				startJ = j
				startJv = endJ
			}
		}
	}
	export.field.StartRows = export.headMaxLevel + 1
	startNowRow := 0
	fieldLen := len(export.allFields)
	f := func(row, cell, endRow int, value interface{}) error {
		cellCode := cellMap[cell]
		rowCode := strconv.Itoa(row)
		startCode := cellCode + rowCode
		style := cellStyleList[cell]
		if row > startNowRow {
			if rowStyle != nil {
				styleStr := rowStyle(row)
				if styleStr != "" {
					if _, ok := hashStyle[styleStr]; !ok {
						hashStyle[styleStr], err = xlsx.NewStyle(styleStr)
						if err != nil {
							return err
						}
					}
					xlsx.SetCellStyle(tableName, cellMap[1]+strconv.Itoa(row),
						cellMap[fieldLen]+strconv.Itoa(row), hashStyle[styleStr])
				}
			}
			startNowRow = row
		}
		endCode := cellCode + strconv.Itoa(endRow-1)
		if endRow-row != 1 {
			xlsx.MergeCell(tableName, startCode, endCode)
		}
		if cellDo != nil {
			styleStr := ""
			styleStr, value = cellDo(cell, value)
			if styleStr != "" {
				style = hashStyle[styleStr]
				if style == 0 {
					hashStyle[styleStr], err = xlsx.NewStyle(styleStr)
					if err != nil {
						return err
					}
					style = hashStyle[styleStr]
				}
				xlsx.SetCellStyle(tableName, startCode, endCode, style)
			}
		} else {
			value = defaultCalValue(export.allFields[cell-1].kind, value)
		}
		xlsx.SetCellValue(tableName, startCode, value)
		return nil
	}
	n := arr.Len()
	rowIndex := 0
	for ; rowIndex < n; rowIndex++ {
		if err := getModelValueForFieldInfo(arr.Index(rowIndex), export.field, f); err != nil {
			return nil, err
		}
	}
	if !haveStyle && rowStyle == nil && cellDo == nil {
		style, err := xlsx.NewStyle(`{"alignment":{"horizontal":"center","vertical":"center"}}`)
		if err == nil {
			xlsx.SetCellStyle(tableName, cellMap[1]+strconv.Itoa(1), cellMap[len(export.allFields)]+strconv.Itoa(export.field.StartRows-1), style)
		}
	}
	return xlsx, err
}

// ListToExcelSheet1BaseToBytes 通过list数据来生成数据格式 通过结构体 tag export:"一级姓名|姓名1,2"
//
// 生成表头和数据格式 通过,来分割表头名和列位置  表头名通过 |来判断层级
//
// rowStyle 对list行处理 cellDo 单元格处理
func ListToExcelSheet1BaseToBytes(list interface{}, rowStyle func(row int) (style string), cellDo func(cell int, value interface{}) (style string, newValue interface{})) ([]byte, error) {
	excel, err := ListToExcelSheet1Base(list, rowStyle, cellDo)
	if err != nil {
		return nil, err
	}
	bf, err := excel.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return bf.Bytes(), nil
}

// ListToExcelSheet1ToBytes 通过list数据来生成数据格式 通过结构体 tag export:"一级姓名|姓名1,2"
//
// 生成表头和数据格式 通过,来分割表头名和列位置  表头名通过 |来判断层级
//
// 数据级转换成excel[]Byte数组
func ListToExcelSheet1ToBytes(list interface{}) ([]byte, error) {
	excel, err := ListToExcelSheet1(list)
	if err != nil {
		return nil, err
	}
	bf, err := excel.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return bf.Bytes(), nil
}

func getSliceBaseType(list interface{}) (base reflect.Type, err error) {
	t := reflect.TypeOf(list)
	if t.Kind() != reflect.Slice {
		return nil, errors.New("list 必须是Slice")
	}
	return t.Elem(), nil
}

func defaultCalValue(kind reflect.Kind, value interface{}) interface{} {
	switch kind {
	case reflect.Slice, reflect.Struct, reflect.Ptr:
		b, _ := json.Marshal(value)
		return string(b)
	}
	return value
}
