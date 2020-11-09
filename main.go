package main

import (
	"fmt"
	"log"
	"os"
	"runtime"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/dcu/pdf"
)

func main() {

	if err := os.RemoveAll("eA"); err != nil {
		// dirty hack for keeping windows console open
		if runtime.GOOS == "windows" {
			fmt.Println(err)
			fmt.Scanf("h")
		}
		log.Fatal(err)
	}

	if err := os.Mkdir("eA", 0755); err != nil {
		// dirty hack for keeping windows console open
		if runtime.GOOS == "windows" {
			fmt.Println(err)
			fmt.Scanf("h")
		}
		log.Fatal(err)
	}

	if err := os.RemoveAll("gA"); err != nil {
		// dirty hack for keeping windows console open
		if runtime.GOOS == "windows" {
			fmt.Println(err)
			fmt.Scanf("h")
		}
		log.Fatal(err)
	}

	if err := os.Mkdir("gA", 0755); err != nil {
		// dirty hack for keeping windows console open
		if runtime.GOOS == "windows" {
			fmt.Println(err)
			fmt.Scanf("h")
		}
		log.Fatal(err)
	}

	if _, err := readPdf(os.Args[1]); err != nil {
		// dirty hack for keeping windows console open
		if runtime.GOOS == "windows" {
			fmt.Println(err)
			fmt.Scanf("h")
		}
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

			kursliste := strings.Contains(row.Content[1].S, "Kursliste")

			if kursliste {
				termline := row.Content[5].S

				// EP or Q?
				ep := strings.Contains(termline, "Einführungsphase")

				// Extract the current term
				// TODO: Better error handling instead of ignoring.
				hj, _ := strconv.Atoi(termline[21:22])

				// Create a new sheet and filename with name of "kurs"
				course := row.Content[9].S
				sheet := course[:strings.IndexByte(course, ' ')]
				filename = course[:strings.IndexByte(course, ' ')]

				x.NewSheet(sheet)

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

				x.SetCellStyle(sheet, "A12", "H12", styleBold)
				// Nr
				x.SetCellValue(sheet, "A12", "Nr.")
				// Name
				x.SetColWidth(sheet, "B", "B", 30)
				x.SetCellValue(sheet, "B12", row.Content[17].S)
				// Hj. 1
				x.SetCellValue(sheet, "C12", row.Content[19].S)
				// Hj. 2
				x.SetCellValue(sheet, "D12", row.Content[21].S)
				if ep {
					// Bemerkungen
					x.SetColWidth(sheet, "E", "E", 30)
					x.SetCellValue(sheet, "E12", row.Content[23].S)
					// Anzahl Fehltage
					x.SetCellValue(sheet, "F12", "Fehltage")
				} else {
					// Hj. 3
					x.SetCellValue(sheet, "E12", row.Content[23].S)
					// Hj. 4
					x.SetCellValue(sheet, "F12", row.Content[25].S)
					// Bemerkungen
					x.SetColWidth(sheet, "G", "G", 30)
					x.SetCellValue(sheet, "G12", row.Content[27].S)
					// Anzahl Fehltage
					x.SetCellValue(sheet, "H12", "Fehltage")
				}

				var j = 0
				for i, word := range row.Content {
					//fmt.Println(word.S)
					//if i < 30 {
					//	continue
					//}
					// Set value of a cell.
					// i> 27 if 4. Halbjahre.
					if i > 24 {
						if len(word.S) > 5 && !strings.Contains(word.S, "Datum,") && !strings.Contains(word.S, "___") {
							j++
							x.SetCellValue(sheet, "A"+fmt.Sprintf("%d", 12+j), j)
							x.SetCellValue(sheet, "B"+fmt.Sprintf("%d", 12+j), word.S)
							//fmt.Println(word.S)
							// TODO: Set older notes
							if hj > 1 {
								x.SetCellValue(sheet, "C"+fmt.Sprintf("%d", 12+j), row.Content[i+4].S)
							}
							if hj > 2 {
								x.SetCellValue(sheet, "D"+fmt.Sprintf("%d", 12+j), row.Content[i+6].S)
							}
							if hj == 4 {
								x.SetCellValue(sheet, "E"+fmt.Sprintf("%d", 12+j), row.Content[i+8].S)
							}
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
