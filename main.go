package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/dcu/pdf"
)

func main() {

	if err := os.Mkdir("eA", 0755); err != nil {
		log.Fatal(err)
	}

	if err := os.Mkdir("gA", 0755); err != nil {
		log.Fatal(err)
	}

	if _, err := readPdf(os.Args[1]); err != nil {
		log.Fatal(err)
	}
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

	for pageIndex := 1; pageIndex <= totalPage; pageIndex++ {
		p := r.Page(pageIndex)
		if p.V.IsNull() {
			continue
		}

		rows, _ := p.GetTextByRow()
		var filename string
		x := excelize.NewFile()

		//style, _ := x.NewStyle(`{"number_format": 1}`)
		styleHeader, _ := x.NewStyle(`{"fill":{"type":"pattern","color":["#FFFFFF"],"pattern":1}}`)
		styleHeaderBold, _ := x.NewStyle(`{"font":{"bold":true,"size":13},"fill":{"type":"pattern","color":["#FFFFFF"],"pattern":1}}`)
		styleBold, _ := x.NewStyle(`{"font":{"bold":true}}`)
		//fmt.Printf("%v", rows)
		for _, row := range rows {
			//println(">>>> row: ", row.Position)
			if row.Content.Len() > 27 {
				// Set value of a cell.
				// Create a new sheet with name of "kurs"
				course := row.Content[9].S
				sheet := course[:strings.IndexByte(course, ' ')]
				filename = course[:strings.IndexByte(course, ' ')]
				x.NewSheet(sheet)
				x.SetColWidth(sheet, "B", "B", 30)
				x.SetColWidth(sheet, "G", "G", 30)

				// Header begin
				x.SetCellStyle(sheet, "A1", "H11", styleHeader)
				// Kursliste
				x.SetCellValue(sheet, "A2", row.Content[1].S)
				x.SetCellStyle(sheet, "A2", "A2", styleHeaderBold)
				// Kurs
				x.SetCellValue(sheet, "C2", row.Content[9].S)
				x.SetCellStyle(sheet, "C2", "C2", styleHeaderBold)
				// Angelaschule Osnabrück
				x.SetCellValue(sheet, "A4", row.Content[3].S)
				// Abiturjahrgang
				x.SetCellValue(sheet, "A6", row.Content[5].S)
				// Kursleiter
				x.SetCellValue(sheet, "A8", row.Content[11].S)
				x.SetCellStyle(sheet, "A8", "A8", styleHeaderBold)
				//Leiter
				x.SetCellValue(sheet, "C8", row.Content[13].S)
				x.SetCellStyle(sheet, "C8", "C8", styleHeaderBold)
				// Zwischenstände
				x.SetCellValue(sheet, "A10", row.Content[15].S)
				// Header end

				x.SetCellStyle(sheet, "B12", "H12", styleBold)
				// Name
				x.SetCellValue(sheet, "B12", row.Content[17].S)
				// Hj. 1
				x.SetCellValue(sheet, "C12", row.Content[19].S)
				// Hj. 2
				x.SetCellValue(sheet, "D12", row.Content[21].S)
				// Hj. 3
				x.SetCellValue(sheet, "E12", row.Content[23].S)
				// Hj. 4
				x.SetCellValue(sheet, "F12", row.Content[25].S)
				// Bemerkungen
				x.SetCellValue(sheet, "G12", row.Content[27].S)
				// Anzahl Fehltage
				x.SetCellValue(sheet, "H12", "Fehltage")

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
						if word.S != "---" {
							x.SetCellValue(sheet, "A"+fmt.Sprintf("%d", 12+j), j)
							x.SetCellValue(sheet, "B"+fmt.Sprintf("%d", 12+j), word.S)
							//fmt.Println(word.S)
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

		if strings.ToUpper(filename) == filename {

			if err := x.SaveAs("eA/" + filename + ".xlsx"); err != nil {
				log.Fatal(err)
			}
		} else {

			if err := x.SaveAs("gA/" + filename + ".xlsx"); err != nil {
				log.Fatal(err)
			}

		}

	}

	return "", nil
}
