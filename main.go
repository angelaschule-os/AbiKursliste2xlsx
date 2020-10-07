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
	return
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
		// Create a new sheet.
		sheet := "Sheet" + fmt.Sprintf("%d", pageIndex)
		x.NewSheet(sheet)
		x.SetColWidth(sheet, "B", "B", 30)
		x.SetColWidth(sheet, "G", "G", 30)

		rows, _ := p.GetTextByRow()
		//fmt.Printf("%v", rows)
		for _, row := range rows {
			println(">>>> row: ", row.Position)
			// Set value of a cell.
			x.SetCellValue(sheet, "B1", "Name")
			x.SetCellValue(sheet, "C1", "Hj. 1")
			x.SetCellValue(sheet, "D1", "Hj. 2")
			x.SetCellValue(sheet, "E1", "Hj. 3")
			x.SetCellValue(sheet, "F1", "Hj. 4")
			x.SetCellValue(sheet, "G1", "Bemerkungen")

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
	// Set active sheet of the workbook.
	x.SetActiveSheet(0)
	// Save xlsx file by the given path.
	if err := x.SaveAs("AbiKursliste.xlsx"); err != nil {
		fmt.Println(err)
	}
	return "", nil
}
