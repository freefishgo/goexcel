package goexcel

import (
	"reflect"
	"sort"
	"strconv"
	"strings"
)

type fieldInfo struct {
	name      string
	headStyle string
	cellStyle string
	headInfo  []string
	kind      reflect.Kind
	sort      int
	index     int
}

type modelInfo struct {
	headMaxLevel int
	field        *newFieldInfo
	allFields    []*fieldInfo
}

type newFieldInfo struct {
	Fields        []*fieldInfo
	StartRows     int
	LevelIndex    int
	LevelIsStruct bool
	Level         *newFieldInfo
}

func getExportSort(src reflect.Type) *modelInfo {
	var list []*fieldInfo
	startSort := 1000
	model := new(modelInfo)
	var f func(reflect.Type, *newFieldInfo)
	f = func(dest reflect.Type, field *newFieldInfo) {
		for dest.Kind() == reflect.Ptr {
			dest = dest.Elem()
		}
		switch dest.Kind() {
		case reflect.Struct:
			for n := 0; n < dest.NumField(); n++ {
				tmpField := field
				tf := dest.Field(n)
				tp := tf.Type
				for tp.Kind() == reflect.Ptr {
					tp = tp.Elem()
				}
				column, ok := tf.Tag.Lookup("export")
				if tp.Kind() == reflect.Slice && !ok {
					tmpField.Level = new(newFieldInfo)
					tmpField.LevelIndex = n
					tp = tp.Elem()
					for tp.Kind() == reflect.Ptr {
						tp = tp.Elem()
					}
					if tp.Kind() == reflect.Struct {
						tmpField.LevelIsStruct = true
						f(tp, tmpField.Level)
						continue
					}
					tmpField = tmpField.Level
				}
				if ok {
					if column == "" {
						column = strings.Split(tf.Tag.Get("json"), ",")[0]
						if column == "" {
							column = tf.Name
						}
					}
					startSort++
					sort := startSort
					name := column
					sp := strings.Split(column, ",")
					if len(sp) > 1 {
						sort, _ = strconv.Atoi(sp[1])
					}
					info := &fieldInfo{
						name:      name,
						sort:      sort,
						index:     n,
						kind:      tf.Type.Kind(),
						headStyle: tf.Tag.Get("headStyle"),
						cellStyle: tf.Tag.Get("cellStyle"),
					}
					if len(sp) > 0 {
						name = sp[0]
						list := strings.Split(name, "|")
						info.name = list[len(list)-1]
						info.headInfo = list
						if model.headMaxLevel < len(list) {
							model.headMaxLevel = len(list)
						}
					}
					list = append(list, info)
					tmpField.Fields = append(tmpField.Fields, info)
				}
			}
		}
	}
	fieldTree := new(newFieldInfo)
	model.field = fieldTree
	fieldTree.StartRows = 1
	f(src, fieldTree)
	sort.Slice(list, func(i, j int) bool {
		if list[i].sort == list[j].sort {
			return list[i].name < list[i].name
		}
		return list[i].sort < list[j].sort
	})
	for k, v := range list {
		v.sort = k + 1
	}
	model.allFields = list
	return model
}

// getModelValueForFieldInfo 获取指定数据实体对应索引的值对指定索引执行 f 方法 f返回值不为空时 会返回
func getModelValueForFieldInfo(dest reflect.Value, field *newFieldInfo, f func(row, cell, endRow int, value interface{}) error) (err error) {
	for dest.Kind() == reflect.Ptr {
		dest = dest.Elem()
	}
	endRows := field.StartRows
	if field.Level != nil {
		field.Level.StartRows = field.StartRows
		list := dest.Field(field.LevelIndex)
		n := list.Len()
		for i := 0; i < n; i++ {
			value := list.Index(i)
			for value.Kind() == reflect.Ptr {
				value = value.Elem()
			}
			if field.LevelIsStruct {
				err = getModelValueForFieldInfo(value, field.Level, f)
				if err != nil {
					return err
				}
				endRows = field.Level.StartRows
			} else {
				err = f(field.Level.StartRows, field.Level.Fields[0].sort, field.Level.StartRows+1, value)
				field.Level.StartRows++
				if err != nil {
					return err
				}
			}
		}
		endRows = field.Level.StartRows
	}
	if endRows == field.StartRows {
		endRows++
	}
	for _, v := range field.Fields {
		value := dest.Field(v.index).Interface()
		err = f(field.StartRows, v.sort, endRows, value)
		if err != nil {
			return err
		}
	}
	field.StartRows = endRows
	return
}

// getModelValueForIndex 获取指定数据实体对应索引的值对指定索引执行 f 方法 f返回值不为空时 会返回
func getModelValueForIndex(dest reflect.Value, indexes [][]int, f func(index int, value interface{}) error) (err error) {
	for dest.Kind() == reflect.Ptr {
		dest = dest.Elem()
	}
cont:
	for i := 0; i < len(indexes); i++ {
		son := indexes[i]
		destSon := dest
		for j := 0; j < len(son); j++ {
			destSon = destSon.Field(son[j])
			if destSon.Kind() == reflect.Ptr {
				if destSon.IsNil() {
					continue cont
				}
				destSon = destSon.Elem()
			}
		}
		if err = f(i, destSon.Interface()); err != nil {
			return err
		}
	}
	return
}
