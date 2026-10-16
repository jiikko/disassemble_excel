package main

// Fixture of the end-to-end test. testdata/sample.xlsm is committed; rebuild it only when the fixture itself should
// change, then refresh the golden output and review the diff of testdata/golden/:
//
//	go test -run TestE2E -regen -update .
//
// vbaProject.bin is excelize's test data (github.com/xuri/excelize/v2 test/vbaProject.bin, BSD-3-Clause): building
// one needs Excel.
//
// excelize cannot write everything Excel writes, so after saving, the XML of a few parts is patched to the form Excel
// produces: an x14 (Excel 2010+) data validation in <extLst> whose list refers to another sheet, a
// hidden defined name, a _xHHHH_ escape in a shared string, and a table calculated-column formula.

import (
	"archive/zip"
	"bytes"
	_ "embed"
	"fmt"
	"io"
	"os"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

//go:embed testdata/vbaProject.bin
var vbaProject []byte

func buildFixture(t *testing.T, path string) {
	t.Helper()
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatal(err)
		}
	}

	must(f.SetSheetName("Sheet1", "Summary"))
	_, err := f.NewSheet("Data")
	must(err)
	_, err = f.NewSheet("Lists")
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

	// Sheet parts besides cells (sheetextras.go).
	set("A10", "エスケープ")
	set("B10", "ESCAPE_PLACEHOLDER") // replaced in the XML below: excelize would protect a literal _x000D_
	must(f.AddComment("Summary", excelize.Comment{Cell: "B7", Author: "Reviewer", Text: "Data の小計の合計"}))
	must(f.SetCellHyperLink("Summary", "A1", "https://example.com/spec", "External"))
	whole := excelize.NewDataValidation(true)
	whole.Sqref = "B2:B5"
	must(whole.SetRange(1, 100, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween))
	must(f.AddDataValidation("Data", whole))
	must(f.AddTable("Data", &excelize.Table{Range: "A1:C5", Name: "Lines"}))
	for i, v := range []string{"東京", "大阪", "福岡"} {
		must(f.SetCellValue("Lists", fmt.Sprintf("A%d", i+1), v))
	}
	must(f.SetSheetVisible("Lists", false, true)) // veryHidden
	must(f.AddVBAProject(vbaProject))
	must(f.AddVBAProject(vbaProject))
	var buf bytes.Buffer
	must(f.Write(&buf))
	must(os.WriteFile(path, patchZip(t, buf.Bytes(), fixturePatches), 0o644))
}

// fixturePatches are exact replacements; each old string must occur exactly once in its part.
var fixturePatches = map[string][][2]string{
	"xl/sharedStrings.xml": {{"ESCAPE_PLACEHOLDER", "CR_x000D_LF and literal _x005F_x0041_"}},
	"xl/workbook.xml":      {{`<definedName name="TaxRate"`, `<definedName name="TaxRate" hidden="1"`}},
	"xl/tables/table1.xml": {{`<tableColumn id="3" name="小計"></tableColumn>`,
		`<tableColumn id="3" name="小計"><calculatedColumnFormula>Lines[[#This Row],[単価]]*Lines[[#This Row],[数量]]</calculatedColumnFormula></tableColumn>`}},
	"xl/worksheets/sheet1.xml": {{"</worksheet>", `<extLst>` +
		`<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
		`<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
		`<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">` +
		`<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C2:C5</xm:sqref></x14:dataValidation>` +
		`</x14:dataValidations></ext>` +
		`</extLst></worksheet>`}},
}

func patchZip(t *testing.T, data []byte, patches map[string][][2]string) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	applied := map[string]bool{}
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		s := string(b)
		for _, p := range patches[f.Name] {
			if n := strings.Count(s, p[0]); n != 1 {
				t.Fatalf("patch %s: %q occurs %d times, want 1", f.Name, p[0], n)
			}
			s = strings.Replace(s, p[0], p[1], 1)
		}
		applied[f.Name] = true
		w, err := zw.Create(f.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := io.WriteString(w, s); err != nil {
			t.Fatal(err)
		}
	}
	for name := range patches {
		if !applied[name] {
			t.Fatalf("patch target %s is not in the workbook", name)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}
