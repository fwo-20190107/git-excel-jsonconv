package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/fwo-20190107/xls"
	"github.com/tealeg/xlsx"
)

func main() {
	if len(os.Args) != 2 {
		log.Fatal("Usage: git-excel-jsonconv file.xls[x]")
	}
	excelFileName := os.Args[1]
	pos := strings.LastIndex(excelFileName, ".")
	ext := excelFileName[pos:]

	var rows [][]string
	if ext == ".xls" {
		xlFile, err := xls.Open(excelFileName, "utf-8")
		if err != nil {
			log.Fatal(err)
		}
		for i := 0; i < xlFile.NumSheets(); i++ {
			if sheet := xlFile.GetSheet(i); sheet != nil {
				for r := 0; r <= int(sheet.MaxRow); r++ {
					row := sheet.Row(r)
					if row == nil {
						continue
					}

					var cols []string
					for c := row.FirstCol(); c <= row.LastCol(); c++ {
						cols = append(cols, row.Col(c))
					}
					rows = append(rows, cols)
				}
			}
		}
	}
	if ext == ".xlsx" {
		xlFile, err := xlsx.OpenFile(excelFileName)
		if err != nil {
			log.Fatal(err)
		}

		for _, sheet := range xlFile.Sheets {
			for _, row := range sheet.Rows {
				if row == nil {
					continue
				}

				cels := make([]string, len(row.Cells))
				for i, cell := range row.Cells {
					var s string
					if cell.Type() == xlsx.CellTypeStringFormula {
						s = cell.Formula()
					} else {
						s = cell.String()
					}
					cels[i] = s
				}
				rows = append(rows, cels)
			}
		}
	}
	output, err := json.Marshal(&rows)
	if err != nil {
		log.Fatal(err)
	}

	var buf bytes.Buffer
	err = json.Indent(&buf, []byte(output), "", "    ")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(buf.String())
}
