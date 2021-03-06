package exceler

import (
	"bytes"
	"fmt"
	"reflect"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/aymerick/raymond"
)

// ExcelReport struct of excel reporting
type ExcelReport struct {
	f         *excelize.File
	sheetName string
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
		cellOldValue, _ := r.f.GetCellValue(r.sheetName, axis)

		if cellValue == "" && cellOldValue == "" {
			continue
		}
		if cellValue == "" && row[colIndex] == "" {
			continue
		}

		formula, _ := r.f.GetCellFormula(r.sheetName, axis)

		if formula != "" {
			continue
		}

		setCellValue(r.sheetName, axis, r.f, cellValue)
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
			if isArray(ctx, prop) {
				rowNumber++
				continue
			}
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
				arrayLength := arr.Len()

				for i := 0; i < arrayLength-1; i++ {
					r.f.DuplicateRow(r.sheetName, rowNumber)
				}

				for i := 0; i < arrayLength; i++ {
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
	if !(strings.Contains(cellTemplate, "{{") && strings.Contains(cellTemplate, "}}")) {
		return cellTemplate, nil
	}
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
	if out == "NaN" {
		return "", nil
	}
	return out, nil
}

// Render render given data to the template
func (r *ExcelReport) Render(ctx map[string]interface{}) {
	mergedCells, _ := r.f.GetMergeCells(r.sheetName)

	r.renderSingleAttribute(ctx)

	// Set nilai pada cell yang digabungkan (merge cell)
	for _, mergedCell := range mergedCells {
		axis := strings.Split(mergedCell[0], ":")[0]
		formula, _ := r.f.GetCellFormula(r.sheetName, axis)
		if formula != "" {
			r.f.SetCellFormula(r.sheetName, axis, formula)
			continue
		}
		cellValue := mergedCell[1]

		prop := getProp(cellValue)
		if prop != "" {
			if !isArray(ctx, prop) {
				cellValue, _ = renderContext(cellValue, ctx)
			}
		}

		setCellValue(r.sheetName, axis, r.f, cellValue)
	}

	r.renderArrayAttribute(ctx)
}

// Save report to file
func (r *ExcelReport) Save(f string) error {
	return r.f.SaveAs(f)
}

// GetBuffer ...
func (r *ExcelReport) GetBuffer() (*bytes.Buffer, error) {
	return r.f.WriteToBuffer()
}

// GetBinary ...
func (r *ExcelReport) GetBinary() ([]byte, error) {
	buff, err := r.GetBuffer()
	if err != nil {
		return nil, err
	}
	return buff.Bytes(), err
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
