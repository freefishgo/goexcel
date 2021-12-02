package goexcel

import (
	"reflect"
)

type InterIter interface {
	Next() bool
	InterKey() reflect.Value
	InterValue() reflect.Value
	Len() int
}

type mapIter[T1 comparable, T2 any] struct {
	m         *ExportMap[T1, T2]
	iterIndex int
	key       T1
	value     T2
}

func (m *mapIter[T1, T2]) Next() bool {
	m.iterIndex++
	if m.iterIndex >= len(m.m.sort) {
		return false
	}
	m.key = m.m.sort[m.iterIndex]
	m.value = m.m.data[m.key]
	return true
}

func (m *mapIter[T1, T2]) Key() T1 {
	return m.key
}

func (m *mapIter[T1, T2]) Value() T2 {
	return m.value
}

func (m *mapIter[T1, T2]) InterKey() reflect.Value {
	return reflect.ValueOf(m.Key())
}
func (m *mapIter[T1, T2]) InterValue() reflect.Value {
	return reflect.ValueOf(m.Value())
}

func (m *mapIter[T1, T2]) Len() int {
	return m.m.Len()
}

type ExportMapper interface {
	MapSortIterInterface() InterIter
}

type ExportMap[T1 comparable, T2 any] struct {
	data map[T1]T2
	sort []T1
}

func (export *ExportMap[T1, T2]) SetValue(key T1, value T2) {
	if _, ok := export.data[key]; !ok {
		export.sort = append(export.sort, key)
	}
	export.data[key] = value
}

func (export *ExportMap[T1, T2]) GetValue(key T1) T2 {
	return export.data[key]
}

func (export *ExportMap[T1, T2]) MapKeys() []T1 {
	return export.sort
}

func (export *ExportMap[T1, T2]) MapSortIter() *mapIter[T1, T2] {
	return &mapIter[T1, T2]{
		m:         export,
		iterIndex: -1,
	}
}

func (export *ExportMap[T1, T2]) Len() int {
	return len(export.sort)
}

func (export *ExportMap[T1, T2]) MapSortIterInterface() InterIter {
	return export.MapSortIter()
}

func NewExportMap[T1 comparable, T2 any]() *ExportMap[T1, T2] {
	export := &ExportMap[T1, T2]{
		data: map[T1]T2{},
	}
	return export
}
