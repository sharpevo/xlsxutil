package excel

import (
	"encoding/csv"
	"github.com/tealeg/xlsx"
	"os"
)

type Separator int

const (
	SEPARATOR_TAB Separator = iota
	SEPARATOR_COMMA
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
		for _, index := range columnIndices {
			r = append(r, row.Cells[index].String())
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
	output := xlsx.NewFile()
	sheet, _ := output.AddSheet(sheetName)
	for _, r := range data {
		row := sheet.AddRow()
		for _, item := range r {
			cell := row.AddCell()
			cell.Value = item
		}
	}
	err = output.Save(fileName)
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
