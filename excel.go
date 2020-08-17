package excelutil

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

type Separator int

type TerminateLoopError struct{}

func (e *TerminateLoopError) Error() string {
	return ""
}

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
	return extractColumnsByIndices(
		fileName, sheetIndex, rowStartsAt, rowEndsAt, columnIndices)
}

func extractColumnsByIndices(
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
	sheet, err := getSheetByIndex(f, sheetIndex)
	if err != nil {
		return data, err
	}
	if rowEndsAt == xlsx.NoRowLimit {
		rowEndsAt = sheet.MaxRow - 1
	}
	if err := sheet.ForEachRow(func(row *xlsx.Row) error {
		index := row.GetCoordinate()
		if index > rowEndsAt {
			return &TerminateLoopError{}
		}
		if index < rowStartsAt {
			return nil
		}
		r := []string{}
		for _, index := range columnIndices {
			cell := row.GetCell(index)
			r = append(r, strings.TrimRight(cell.String(), "\r\n"))
		}
		data = append(data, r)
		return nil
	}); err != nil {
		if _, ok := err.(*TerminateLoopError); !ok {
			return data, err
		}
	}
	return data, nil
}

func ExtractColumnsByIds(
	fileName string,
	sheetIndex int,
	rowStartsAt int,
	rowEndsAt int,
	columnIds []string,
) (data [][]string, err error) {
	columnIndices := []int{}
	for _, coord := range columnIds {
		columnIndex, _, err := xlsx.GetCoordsFromCellIDString(
			fmt.Sprintf("%s1", coord))
		if err != nil {
			return data, err
		}
		columnIndices = append(columnIndices, columnIndex)
	}
	return extractColumnsByIndices(
		fileName, sheetIndex, rowStartsAt, rowEndsAt, columnIndices)
}

func getSheetByIndex(file *xlsx.File, index int) (*xlsx.Sheet, error) {
	maxSheet := len(file.Sheets)
	if index > maxSheet-1 {
		return nil, fmt.Errorf(
			"sheet index out of range [%d] with lengith %d", index, maxSheet)
	}
	return file.Sheets[index], nil

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
