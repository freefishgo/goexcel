package goexcel

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strings"
)

type excelToList struct {
	excel         *excelize.File
	sheet         string
	model         reflect.Value
	export        *modelInfo
	base          reflect.Type
	rows          [][]string
	mergeCellsMap map[int]map[string]int
	startRow      int
	notSpace      bool //是否不清除空格
}

// setBaseModel 设置基本表结构
func (e *excelToList) setBaseModel(base reflect.Type) (err error) {
	e.export = getExportSort(base, false)
	e.base = base
	e.startRow = e.export.headMaxLevel
	if len(e.export.allFields) == 0 {
		return errors.New("结构体没有导出的字段")
	}
	return
}

func (e *excelToList) setExport(excel *excelize.File) {
	e.excel = excel
	e.sheet = "Sheet1"
	e.rows = append([][]string{{}}, e.excel.GetRows(e.sheet)...)
	for _, v := range e.excel.GetMergeCells(e.sheet) {
		axis := strings.Split(v[0], ":")
		if ok, cell, row := e.export.IsCellFlags(axis[0]); ok {
			if _, ok := e.mergeCellsMap[row]; !ok {
				e.mergeCellsMap[row] = map[string]int{}
			}
			_, row1 := AxisToCellRow(axis[1])
			e.mergeCellsMap[row][cell] = row1
		}
	}
}

// Next 是否有下一行
func (e *excelToList) Next() (have bool) {
	e.startRow++
	if e.startRow >= len(e.rows) {
		return false
	}
	value := reflect.New(e.base).Elem()
	e.startRow = e.formatting(e.export.field, e.startRow, value)
	e.model = value
	return true
}

// GetRow 获取下一行 先调用 Next
func (e *excelToList) GetRow() interface{} {
	return e.model.Interface()
}

// formatting 格式化一条数据
func (e *excelToList) formatting(field *newFieldInfo, startRow int, value reflect.Value) (endRow int) {
	value = getBaseValue(value)
	for _, v := range field.Fields {
		row := e.rows[startRow]
		if len(row) >= v.sort {
			val := row[v.sort-1]
			if !e.notSpace {
				val = strings.TrimSpace(val)
			}
			stringSetValue(value.Field(v.index), val)
		}
	}
	endRow = startRow
	if v, ok := e.mergeCellsMap[startRow]; ok {
		if v, ok := v[field.CellFlag]; ok {
			endRow = v
		}
	}
	if field.Level != nil {
		oldValue := value.Field(field.LevelIndex)
		sliceValue := getBaseValue(oldValue)
		nextRow := startRow
		for ; nextRow <= endRow; nextRow++ {
			if field.LevelIsStruct {
				tv := reflect.New(field.LevelType)
				nextRow = e.formatting(field.Level, nextRow, tv)
				sliceValue = reflect.Append(sliceValue, tv.Elem())
			}
		}
		oldValue.Set(sliceValue)
	}
	return
}

func getBaseValue(value reflect.Value) reflect.Value {
	for value.Type().Kind() == reflect.Ptr {
		if value.IsNil() {
			value.Set(reflect.New(value.Type().Elem()))
		}
		value = value.Elem()
	}
	return value
}

func NewExcelSheet1ToList(excel *excelize.File, model interface{}) (e *excelToList, err error) {
	tp := reflect.TypeOf(model)
	e = new(excelToList)
	e.mergeCellsMap = map[int]map[string]int{}
	err = e.setBaseModel(tp)
	e.setExport(excel)
	return
}

// NewExcelSheet1ToListFromPath path路径excel文件数据转换成 e 结构数据 切片元素export数据必须与
//
// excel中表头和结构一致
func NewExcelSheet1ToListFromPath(path string, model interface{}) (e *excelToList, err error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return
	}
	return NewExcelSheet1ToList(f, model)
}

// ExcelSheet1ToListFromPath path路径excel文件数据转换成 listPtr 结构数据 切片元素export数据必须与
//
// excel中表头和结构一致
func ExcelSheet1ToListFromPath(path string, listPtr interface{}) (err error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return
	}
	return ExcelSheet1ToList(f, listPtr)
}

// ExcelSheet1ToList excel文件数据转换成 listPtr 结构数据 切片元素export数据必须与
//
// excel中表头和结构一致
func ExcelSheet1ToList(excel *excelize.File, listPtr interface{}) (err error) {
	base, err := getPtrSliceBaseType(listPtr)
	if err != nil {
		return err
	}
	ex, err := NewExcelSheet1ToList(excel, reflect.New(base).Elem().Interface())
	if err != nil {
		return err
	}
	v := reflect.ValueOf(listPtr)
	slice := getBaseValue(v)
	for ex.Next() {
		slice = reflect.Append(slice, ex.model)
	}
	for v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	v.Set(slice)
	return
}

func getPtrSliceBaseType(list interface{}) (base reflect.Type, err error) {
	t := reflect.TypeOf(list)
	if t.Kind() != reflect.Ptr {
		return nil, errors.New("list 必须是Slice地址")
	}
	for t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Slice {
		return nil, errors.New("list 必须是Slice地址")
	}
	return t.Elem(), nil
}
