// Command gen writes testdata/sample.xlsm, the fixture of the end-to-end test (e2e_test.go).
//
//	go run ./testdata/gen
//
// The fixture is committed, so this only needs to run when the fixture itself should change.
// After regenerating, refresh the golden output with `go test -run TestE2E -update .` and
// review the diff of testdata/golden/.
//
// vbaProject.bin is copied from excelize's test data (github.com/xuri/excelize/v2 test/vbaProject.bin,
// BSD-3-Clause). Building one from scratch needs Excel, so the e2e fixture reuses it.
package main

import (
	_ "embed"
	"log"
	"time"

	"github.com/xuri/excelize/v2"
)

//go:embed vbaProject.bin
var vbaProject []byte

func main() {
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	must(f.SetSheetName("Sheet1", "Summary"))
	_, err := f.NewSheet("Data")
	must(err)

	// Summary: value types the cells.tsv / values.csv must keep apart.
	set := func(cell string, v any) { must(f.SetCellValue("Summary", cell, v)) }
	set("A1", "項目")
	set("B1", "値")
	set("A2", "整数")
	set("B2", 42)
	set("A3", "小数")
	set("B3", 0.125)
	set("A4", "日付")
	set("B4", time.Date(2026, 1, 31, 0, 0, 0, 0, time.UTC))
	set("A5", "真偽値")
	set("B5", true)
	set("A6", "改行とタブ")
	set("B6", "1行目\n2行目\tタブ")
	set("A7", "合計 (別シート参照)")
	must(f.SetCellFormula("Summary", "B7", "SUM(Data!C2:C5)"))
	set("A8", "エラー")
	must(f.SetCellFormula("Summary", "B8", "1/0"))
	set("A9", "名前付き範囲")
	must(f.SetCellFormula("Summary", "B9", "TaxRate*B2"))

	// Data: a shared formula. Only C2 carries the formula text; C3:C5 are children that the tool must expand.
	for i, row := range [][]any{{"単価", "数量", "小計"}, {100, 2}, {250, 4}, {80, 10}, {1200, 1}} {
		cell, err := excelize.CoordinatesToCellName(1, i+1)
		must(err)
		must(f.SetSheetRow("Data", cell, &row))
	}
	shared, ref := excelize.STCellFormulaTypeShared, "C2:C5"
	must(f.SetCellFormula("Data", "C2", "A2*B2", excelize.FormulaOpts{Type: &shared, Ref: &ref}))
	// A shared formula with sheet-qualified references: children must move Summary!B2 too (issue #2).
	must(f.SetCellValue("Data", "D1", "別シート参照"))
	ref2 := "D2:D5"
	must(f.SetCellFormula("Data", "D2", "Summary!B2+'Summary'!$B$2+C2", excelize.FormulaOpts{Type: &shared, Ref: &ref2}))

	must(f.SetDefinedName(&excelize.DefinedName{Name: "TaxRate", RefersTo: "0.1"}))
	must(f.SetDefinedName(&excelize.DefinedName{Name: "Items", RefersTo: "Data!$A$2:$C$5", Scope: "Data"}))

	// Objects with an assigned macro: a DrawingML shape and a VML form-control button.
	must(f.AddShape("Summary", &excelize.Shape{Cell: "D2", Type: "rect", Macro: "Button1_Click", Width: 120, Height: 40}))
	must(f.AddFormControl("Summary", excelize.FormControl{Cell: "D6", Type: excelize.FormControlButton, Macro: "Button2_Click", Text: "実行", Width: 120, Height: 30}))

	must(f.AddVBAProject(vbaProject))
	must(f.SaveAs("testdata/sample.xlsm"))
}

func must(err error) {
	if err != nil {
		log.Fatal(err)
	}
}
