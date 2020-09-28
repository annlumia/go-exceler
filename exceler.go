package exceler

import (
	"bytes"
	"fmt"
	"reflect"
	"regexp"
	"strconv"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/aymerick/raymond"
)

var (
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rangeRgx    = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
)

// ExcelReport struct of excel reporting
type ExcelReport struct {
	f         *excelize.File
	sheetName string
}

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

func (r *ExcelReport) renderRow(row []string, rowNumber int, ctx map[string]interface{}) {
	renderedRow := make([]string, len(row))

	for columnIndex, columnValue := range row {
		renderedRow[columnIndex], _ = renderContext(columnValue, ctx)
	}

	for colIndex, cellValue := range renderedRow {
		colNum := colIndex + 1
		colName, err := excelize.ColumnNumberToName(colNum)
		if err != nil {
			return
		}

		axis := fmt.Sprintf("%s%v", colName, rowNumber)
		formula, _ := r.f.GetCellFormula(r.sheetName, axis)

		if formula != "" {
			continue
		}
		v, err := strconv.ParseFloat(cellValue, 64)
		if err == nil {
			r.f.SetCellValue(r.sheetName, axis, v)
		} else {
			r.f.SetCellValue(r.sheetName, axis, cellValue)
		}
	}

}

func (r *ExcelReport) renderSingleAttribute(ctx map[string]interface{}) {
	rows, err := r.f.Rows(r.sheetName)
	if err != nil {
		return
	}
	rowNumber := 1

	for rows.Next() {
		row, _ := rows.Columns()

		prop := getListProp(row)
		if prop != "" {
			rowNumber++
			continue
		}

		r.renderRow(row, rowNumber, ctx)

		rowNumber++
	}

}

func (r *ExcelReport) renderArrayAttribute(ctx map[string]interface{}) {
	rows, err := r.f.Rows(r.sheetName)
	if err != nil {
		return
	}

	rowNumber := 1
	for rows.Next() {
		row, _ := rows.Columns()
		prop := getListProp(row)
		if prop != "" {
			if isArray(ctx, prop) {

				arr := reflect.ValueOf(ctx[prop])
				ctxBackup := ctx[prop]

				for i := 0; i < arr.Len()-1; i++ {
					r.f.DuplicateRow(r.sheetName, rowNumber)
				}

				for i := 0; i < arr.Len(); i++ {
					ctx[prop] = arr.Index(i).Interface()
					r.renderRow(row, rowNumber+i, ctx)
				}

				ctx[prop] = ctxBackup
			}
			rowNumber++
			continue
		}

		rowNumber++
	}

}

func renderContext(cellTemplate string, ctx interface{}) (string, error) {
	tpl := strings.Replace(cellTemplate, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
	template, err := raymond.Parse(tpl)
	if err != nil {
		return cellTemplate, err
	}
	out, err := template.Exec(ctx)
	if err != nil {
		return cellTemplate, err
	}
	return out, nil
}

// Render render given data to the template
func (r *ExcelReport) Render(ctx map[string]interface{}) {
	mergedCells, _ := r.f.GetMergeCells(r.sheetName)

	r.renderSingleAttribute(ctx)
	r.renderArrayAttribute(ctx)

	// Set nilai pada cell yang digabungkan (merge cell)
	for _, mergedCell := range mergedCells {
		axis := strings.Split(mergedCell[0], ":")[0]
		formula, _ := r.f.GetCellFormula(r.sheetName, axis)
		if formula != "" {
			r.f.SetCellFormula(r.sheetName, axis, formula)
			continue
		}
		cellValue := mergedCell[1]
		cellValue, _ = renderContext(cellValue, ctx)

		v, err := strconv.ParseFloat(cellValue, 64)
		if err == nil {
			r.f.SetCellValue(r.sheetName, axis, v)
		} else {
			r.f.SetCellValue(r.sheetName, axis, cellValue)
		}
	}

}

// Save report to file
func (r *ExcelReport) Save(f string) error {
	return r.f.SaveAs(f)
}

// NewFromFile Create new template from excel file
func NewFromFile(filename string, sheetName string) (*ExcelReport, error) {
	template, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}

	return &ExcelReport{
		f:         template,
		sheetName: sheetName,
	}, nil
}

// NewFromBinary Create new template from binary
func NewFromBinary(b []byte, sheetName string) (*ExcelReport, error) {
	r := bytes.NewReader(b)

	template, err := excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}

	return &ExcelReport{
		f:         template,
		sheetName: sheetName,
	}, nil
}
