package main

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/dcu/pdf"
)

func main() {
	content, err := readPdf(os.Args[1]) // Read local pdf file
	if err != nil {
		panic(err)
	}
	fmt.Println(content)
}

func readPdf(path string) (string, error) {
	f, r, err := pdf.Open(path)
	defer func() {
		_ = f.Close()
	}()
	if err != nil {
		return "", err
	}
	totalPage := r.NumPage()

	x := excelize.NewFile()

	for pageIndex := 1; pageIndex <= totalPage; pageIndex++ {
		p := r.Page(pageIndex)
		if p.V.IsNull() {
			continue
		}

		rows, _ := p.GetTextByRow()
		//fmt.Printf("%v", rows)
		for _, row := range rows {
			println(">>>> row: ", row.Position)
			if row.Content.Len() > 27 {
				// Set value of a cell.
				// Create a new sheet with name of "kurs"
				sheet := row.Content[9].S
				x.NewSheet(sheet)
				x.SetColWidth(sheet, "B", "B", 30)
				x.SetColWidth(sheet, "G", "G", 30)
				// Name
				x.SetCellValue(sheet, "B1", row.Content[17].S)
				// Hj. 1
				x.SetCellValue(sheet, "C1", row.Content[19].S)
				// Hj. 2
				x.SetCellValue(sheet, "D1", row.Content[21].S)
				// Hj. 3
				x.SetCellValue(sheet, "E1", row.Content[23].S)
				// Hj. 4
				x.SetCellValue(sheet, "F1", row.Content[25].S)
				// Bemerkungen
				x.SetCellValue(sheet, "G1", row.Content[27].S)

				var i = 0
				var j = 1
				for _, word := range row.Content {
					//fmt.Println(word.S)
					i++
					//if i < 30 {
					//	continue
					//}
					// Set value of a cell.
					if i > 27 && i%4 != 0 && i%2 == 0 {
						x.SetCellValue(sheet, "A"+fmt.Sprintf("%d", 1+j), j)
						x.SetCellValue(sheet, "B"+fmt.Sprintf("%d", 1+j), word.S)
						fmt.Println(word.S)
						j++
					}
				}
			}
		}
	}
	// Delete standard sheet
	x.DeleteSheet("Sheet1")
	// Set active sheet of the workbook.
	x.SetActiveSheet(2)
	// Save xlsx file by the given path.
	if err := x.SaveAs("AbiKursliste.xlsx"); err != nil {
		fmt.Println(err)
	}
	return "", nil
}
