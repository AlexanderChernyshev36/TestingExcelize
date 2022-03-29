package main

import (
"fmt"
"strconv"

"github.com/xuri/excelize"
)

func main() {

	f := excelize.NewFile()

	nameSheet_1 := "НовыйЛист"
	f.SetSheetName("Sheet1", nameSheet_1)
	//nameSheet_2 := "НовыйЛист2"
	//f.NewSheet(nameSheet_2)

	StyleBorderBlack, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
		},
	})

	for i := 1; i <= 10; i++{
		f.SetCellValue(nameSheet_1, "A"+strconv.Itoa(i),i)
	}

	for i := 1; i <= 10; i++{
		f.MergeCell(nameSheet_1, "B"+strconv.Itoa(i),"C" + strconv.Itoa(i))
	}

	_ = f.SetCellStyle(nameSheet_1, "A1", "A10", StyleBorderBlack)

	_ = f.SetCellStyle(nameSheet_1, "D1", "D10", StyleBorderBlack)

	_ = f.SetCellStyle(nameSheet_1, "B1", "B10", StyleBorderBlack)

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

