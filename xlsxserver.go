package main

import (
	"fmt"
	"strconv"

	"github.com/xlsx"
)

type KQreport struct {
	result []string
	index  int
}

func readxlsx(url string) []KQreport {
	reports := make([]KQreport, 0, 1000)

	xlFile, err := xlsx.OpenFile(url)
	if err != nil {
		fmt.Printf("Open failed: %s\n", err)
	}
	for _, sheet := range xlFile.Sheets {
		index := 0

		for _, row := range sheet.Rows {
			if index > 0 {
				var results []string = make([]string, 0, 10)
				for _, cell := range row.Cells {
					text := cell.String()
					results = append(results, text)
				}
				fmt.Printf("%s\n", results)
				temp := KQreport{results, index}
				reports = append(reports, temp)
			}
			index++

		}
	}
	return reports
}

func writeXlsx(url string, reports []KQreport) {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	var name string
	if len(url) > 1 {
		name = url
	} else {
		name = "save.xlsx"
	}
	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	for _, item := range reports {
		row = sheet.AddRow()
		for _, sa := range item.result {
			cell = row.AddCell()
			cell.Value = sa
		}
	}

	err = file.Save(name)
	if err != nil {
		fmt.Printf(err.Error())
	}
	// row = sheet.AddRow()
	// cell = row.AddCell()
	// cell.Value = "I am a cell!"

}

func translateXl(i int, reports KQreport) []KQreport {
	var index = i
	indes := strconv.Itoa(index)
	z_reports := make([]KQreport, 0, 1000)
	// name := reports.result[0]
	name := reports.result[1]
	// dep := reports.result[2]
	lottery1 := reports.result[5]
	lottery2 := reports.result[6]
	lottery3 := reports.result[7]

	l1, _ := strconv.Atoi(lottery1)
	l2, _ := strconv.Atoi(lottery2)
	l3, _ := strconv.Atoi(lottery3)
	count := 0
	for i := 0; i < l1; i++ {
		cell := []string{indes, name + strconv.Itoa(count), "", "1", ""}
		fmt.Println(cell)
		report := KQreport{cell, index}
		z_reports = append(z_reports, report)
		count++
	}

	for i := 0; i < l2; i++ {
		cell := []string{indes, name + strconv.Itoa(count), "", "2", ""}
		report := KQreport{cell, index}
		z_reports = append(z_reports, report)
		count++
	}

	for i := 0; i < l3; i++ {
		cell := []string{indes, name + strconv.Itoa(count), "", "3", ""}
		report := KQreport{cell, index}
		z_reports = append(z_reports, report)
		count++
	}

	return z_reports
}
