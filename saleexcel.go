package main

func main() {
	var reports = readxlsx("test.xlsx")
	fin_reports := translate(reports)
	writeXlsx("finalss.xlsx", fin_reports)
}

func translate(reports []KQreport) []KQreport {
	var fin_reports []KQreport
	for _, report := range reports {
		zreport := translateXl(1, report)
		for _, treport := range zreport {
			fin_reports = append(fin_reports, treport)
		}
	}
	return fin_reports
}
