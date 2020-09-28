package exceler

import "reflect"

func getRangeProp(in []string) string {
	for _, cellValue := range in {
		match := rangeRgx.FindAllStringSubmatch(cellValue, -1)
		if match != nil {
			return match[0][1]
		}
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
