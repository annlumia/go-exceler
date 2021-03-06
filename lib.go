package exceler

import (
	"reflect"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func getRangeProp(in []string) string {
	for _, cellValue := range in {
		match := rangeRgx.FindAllStringSubmatch(cellValue, -1)
		if match != nil {
			return match[0][1]
		}
	}
	return ""
}

func getProp(in string) string {
	if match := rgx.FindAllStringSubmatch(in, -1); match != nil {
		return match[0][1]
	}
	return ""
}

func getListProp(in []string) string {
	for _, cell := range in {
		if match := rgx.FindAllStringSubmatch(cell, -1); match != nil {
			return match[0][1]
		}
	}
	return ""
}

func isArray(in map[string]interface{}, prop string) bool {
	val, ok := in[prop]
	if !ok {
		return false
	}
	switch reflect.TypeOf(val).Kind() {
	case reflect.Array, reflect.Slice:
		return true
	}
	return false
}

func setCellValue(sheet string, axis string, f *excelize.File, value string) {
	v, err := strconv.ParseFloat(value, 64)
	if err == nil {
		f.SetCellValue(sheet, axis, v)
		return
	}

	if value != value {
		// value is NaN
		f.SetCellValue(sheet, axis, "")
		return
	}

	f.SetCellValue(sheet, axis, value)
}
