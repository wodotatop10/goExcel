package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/xlsx"
)

type KQreport struct {
	result []string
	index  int
}

const MAXLINE int = 2000
const MAXCELL int = 20

func readxlsx(url string) []KQreport {
	reports := make([]KQreport, 0, MAXLINE)

	xlFile, err := xlsx.OpenFile(url)
	if err != nil {
		fmt.Printf("Open failed: %s\n", err)
	}
	for _, sheet := range xlFile.Sheets {
		index := 0

		for _, row := range sheet.Rows {
			if index > 0 {
				var results []string = make([]string, 0, MAXCELL)
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
		name = "save"
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

	d := 1
	// isfile, err := PathExists(name + ".xlsx")
	// if isfile {
	// 	name = name + strconv.Itoa(d)
	// }

	usfile(&name, &d)

	err = file.Save(name + ".xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
	// row = sheet.AddRow()
	// cell = row.AddCell()
	// cell.Value = "I am a cell!"

}

func usfile(name *string, d *int) {
	isfile, _ := PathExists(*name + ".xlsx")
	if isfile {
		*name += strconv.Itoa(*d)
		*d++
		usfile(name, d)
	}
}

func translateXl(i int, reports KQreport) ([]KQreport, int) {
	var index = i
	// indes := strconv.Itoa(index)
	z_reports := make([]KQreport, 0, MAXLINE)
	// name := reports.result[0]
	name := reports.result[1]
	sim := reports.result[2]
	dep := reports.result[3]

	lotterys := []string{reports.result[5], reports.result[6], reports.result[7]}

	count := 0
	for tag, value := range lotterys {
		tagss := strconv.Itoa(tag + 1)
		ls, _ := strconv.Atoi(value)
		for i := 0; i < ls; i++ {
			cell := []string{strconv.Itoa(index), name + strconv.Itoa(count+1), sim, "", tagss, "", dep}
			fmt.Println(cell)
			report := KQreport{cell, index}
			z_reports = append(z_reports, report)
			count++
			index++
		}
	}
	return z_reports, index
}

func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}
