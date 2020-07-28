package excel

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

type Separator int

const (
	SEPARATOR_TAB Separator = iota
	SEPARATOR_COMMA
)

const (
	OUTPUT_TYPE_TXT  = "txt"
	OUTPUT_TYPE_CSV  = "csv"
	OUTPUT_TYPE_XLSX = "xlsx"
)

func ExtractColumns(
	fileName string,
	sheetIndex int,
	rowStartsAt int,
	rowEndsAt int,
	columnIndices []int,
) (data [][]string, err error) {
	f, err := xlsx.OpenFile(fileName)
	if err != nil {
		return data, err
	}
	if sheetIndex > len(f.Sheets) {
		return data, fmt.Errorf(
			"sheet index '%v' out of bounds '%v'", sheetIndex, len(f.Sheets))
	}
	sheet := f.Sheets[sheetIndex]
	for index, row := range sheet.Rows {
		if rowEndsAt != -1 &&
			index > rowEndsAt {
			break
		}
		if index < rowStartsAt ||
			row.Cells[1].String() == "" {
			continue
		}
		r := []string{}
		cellAmount := len(row.Cells)
		for _, index := range columnIndices {
			if index > cellAmount {
				return data, fmt.Errorf(
					"cell index '%v' out of bounds '%v'", index, cellAmount)
			}
			str := row.Cells[index].String()
			r = append(r, strings.TrimRight(str, "\r\n"))
		}
		data = append(data, r)
	}
	return data, nil
}

func MakeFileXLSX(
	fileName string,
	data [][]string,
	sheetName string,
) (err error) {
	file := xlsx.NewFile()
	sheet, _ := file.AddSheet(sheetName)
	for _, r := range data {
		row := sheet.AddRow()
		for _, item := range r {
			cell := row.AddCell()
			cell.Value = item
		}
	}
	err = file.Save(fileName)
	if err != nil {
		return err
	}
	return nil
}

func MakeFileCSV(
	fileName string,
	data [][]string,
	separator Separator,
) (err error) {
	file, err := os.Create(fileName)
	if err != nil {
		return err
	}
	defer file.Close()
	writer := csv.NewWriter(file)
	switch separator {
	case SEPARATOR_TAB:
		writer.Comma = '\t'
	case SEPARATOR_COMMA:
		writer.Comma = ','
	}
	defer writer.Flush()
	for _, items := range data {
		err = writer.Write(items)
		if err != nil {
			return err
		}
	}
	return nil
}
