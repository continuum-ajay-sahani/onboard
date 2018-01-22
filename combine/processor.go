package combine

import (
	"fmt"

	"github.com/ContinuumLLC/onboarding/constdata"
	"github.com/tealeg/xlsx"
)

var regEndpointMap map[string]string
var outputrow [][]string

//ProcessCombine combine data
func ProcessCombine() {
	//combineData()
	getSiteCount()
}

func getSiteCount() {
	excelFileName := "data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println(err)
	}
	siteNameCount := make(map[string]int)
	for index, sheet := range xlFile.Sheets {
		if index != 5 {
			continue
		}
		for rowIndex, row := range sheet.Rows {
			if rowIndex >= constdata.TotalRowSize {
				break
			}
			if rowIndex == 0 {
				// Ignore header row
				continue
			}
			for cellIndex, cell := range row.Cells {
				if cellIndex == 4 {
					siteName := cell.String()
					value, _ := siteNameCount[siteName]
					value++
					siteNameCount[siteName] = value
				}
			}
		}
	}
	fmt.Println(siteNameCount)
}

func getLegacyRegID() {
	excelFileName := "data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println(err)
	}
	//fmt.Println("file open success")
	for index, sheet := range xlFile.Sheets {
		if index != constdata.DataExcelTabIndex {
			continue
		}
		for rowIndex, row := range sheet.Rows {
			if rowIndex >= constdata.TotalRowSize {
				break
			}
			if rowIndex == 0 {
				// Ignore header row
				continue
			}
			for cellIndex, cell := range row.Cells {
				if cellIndex == 3 {
					text := cell.String()
					fmt.Print(text)
					fmt.Print(",")
				}
			}
		}
	}
}

func combineData() {
	createEndpointMap()
	excelFileName := "data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	outputrow = append(outputrow, getHeaderRow())
	for index, sheet := range xlFile.Sheets {
		if index != 5 {
			continue
		}
		for rowIndex, row := range sheet.Rows {
			if rowIndex >= constdata.TotalRowSize {
				break
			}
			if rowIndex == 0 {
				// Ignore header row
				continue
			}
			var endpoint string
			var rowData []string
			for cellIndex, cell := range row.Cells {
				cellData := cell.String()
				if cellIndex == 6 {
					endpoint = regEndpointMap[cellData]
				}
				rowData = append(rowData, cellData)
			}
			rowData = append(rowData, endpoint)
			outputrow = append(outputrow, rowData)
		}
	}
	createFinalSheet()
}

func createFinalSheet() {
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	for _, dataRow := range outputrow {
		row = sheet.AddRow()
		for _, cellData := range dataRow {
			cell = row.AddCell()
			cell.Value = cellData
		}
	}
	err = file.Save("datafinal.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}

func getHeaderRow() []string {
	var headerRow []string
	headerRow = append(headerRow, "MemberID")
	headerRow = append(headerRow, "MemberName")
	headerRow = append(headerRow, "MemberCode")
	headerRow = append(headerRow, "SiteId")
	headerRow = append(headerRow, "SiteName")
	headerRow = append(headerRow, "SiteCode")
	headerRow = append(headerRow, "RegId")
	headerRow = append(headerRow, "ResourceName")
	headerRow = append(headerRow, "RegType")
	headerRow = append(headerRow, "endpointid")
	return headerRow
}

func createEndpointMap() {
	excelFileName := "data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println(err)
	}
	regEndpointMap = make(map[string]string)
	for index, sheet := range xlFile.Sheets {
		if index != constdata.DataExcelTabIndex {
			continue
		}
		for rowIndex, row := range sheet.Rows {
			if rowIndex >= constdata.TotalRowSize {
				break
			}
			if rowIndex == 0 {
				// Ignore header row
				continue
			}
			var endpoint string
			var regID string
			for cellIndex, cell := range row.Cells {
				text := cell.String()
				if cellIndex == 0 {
					endpoint = text
				}
				if cellIndex == 3 {
					regID = text
				}
			}
			regEndpointMap[regID] = endpoint
		}
	}
	//fmt.Println(regEndpointMap)
}
