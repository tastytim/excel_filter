package main

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("test.xlsx")
	if err != nil {
		fmt.Println("errore open", err)
		return
	}

	defer func() {
		if err := f.SaveAs("test1.xlsx"); err != nil {
			fmt.Println("errore.saveas")
		}
		if err := f.Close(); err != nil {
			fmt.Println("close")
		}
	}()

	rows, err := f.GetRows("ALIMENTI_TAVOLE", excelize.Options{})
	if err != nil {
		fmt.Println("getRows")
		return
	}

	lenghtrow := len(rows)

	rowIndex := 1

	for _, row := range rows {
		if row[1] == "ALI_TAVOLE" {
			rowIndex++
			continue
		}

		nums := strings.Split(row[1], "§")

		var nonEmptyValues []string
		for _, value := range nums {
			if value != "" {
				nonEmptyValues = append(nonEmptyValues, value)
			}
		}

		if len(nonEmptyValues) > 1 {
			r := row[0]

			fmt.Println(row[0])

			for _, n := range nonEmptyValues {
				lenghtrow++
				err := f.InsertRows("ALIMENTI_TAVOLE", lenghtrow, 1)
				if err != nil {
					fmt.Println("errore inser row")
				}
				err = f.SetCellValue("ALIMENTI_TAVOLE", fmt.Sprintf("A%d", lenghtrow), r)
				if err != nil {
					fmt.Printf("set cell a %s", err)
				}

				err = f.SetCellValue("ALIMENTI_TAVOLE", fmt.Sprintf("B%d", lenghtrow), regexp.MustCompile(`#?(\d+)#`).FindStringSubmatch(n)[1])
				if err != nil {
					fmt.Printf("set cell b %s", err)
				}
			}

			err := f.RemoveRow("ALIMENTI_TAVOLE", rowIndex)
			if err != nil {
				fmt.Println("errore.removerow")
			}

			lenghtrow--

		} else {
			rowIndex++
		}
	}
}
