package main

func main() {
	var reports = readxlsx("test.xlsx")
	fin_reports := translate(reports)
	writeXlsx("fi", fin_reports)
}

func translate(reports []KQreport) []KQreport {
	var fin_reports []KQreport
	index := 1
	for _, report := range reports {
		zreport, i := translateXl(index, report)
		index = i
		for _, treport := range zreport {
			fin_reports = append(fin_reports, treport)
		}
	}
	return fin_reports
}
