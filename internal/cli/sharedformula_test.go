package cli

import "testing"

func TestShiftFormula(t *testing.T) {
	cases := []struct {
		name       string
		formula    string
		dRow, dCol int
		want       string
	}{
		// issue #2: excelize left sheet-qualified references unshifted
		{"sheet-qualified relative", "Inputs!A1+A1+$A$1+Inputs!$A1", 2, 0, "Inputs!A3+A3+$A$1+Inputs!$A3"},
		{"quoted sheet name", "'My Sheet'!B2*2", 1, 1, "'My Sheet'!C3*2"},
		{"quoted sheet name with escaped quote", "'O''Brien'!A1", 1, 0, "'O''Brien'!A2"},
		{"non-ASCII sheet name", "集計!A1+A1", 1, 0, "集計!A2+A2"},
		{"external workbook", "[1]Sheet1!A1", 0, 2, "[1]Sheet1!C1"},
		{"3D reference", "SUM(Sheet1:Sheet3!A1)", 1, 0, "SUM(Sheet1:Sheet3!A2)"},
		{"range with sheet", "SUM(Data!A1:B2)", 1, 0, "SUM(Data!A2:B3)"},
		{"mixed absolute", "A$1+$A1+$A$1", 3, 3, "D$1+$A4+$A$1"},
		{"whole column", "SUM(A:B)+SUM($C:$C)", 5, 1, "SUM(B:C)+SUM($C:$C)"},
		{"whole row", "SUM(1:2)+SUM($3:$3)", 1, 5, "SUM(2:3)+SUM($3:$3)"},
		{"string literal untouched", `IF(A1="A1",B1,"x""C1")`, 1, 0, `IF(A2="A1",B2,"x""C1")`},
		{"function name untouched", "LOG10(A1)+ATAN2(B1,C1)", 1, 0, "LOG10(A2)+ATAN2(B2,C2)"},
		{"prefixed function untouched", "_xlfn.XLOOKUP(A1,B:B,C:C)", 1, 0, "_xlfn.XLOOKUP(A2,B:B,C:C)"},
		{"structured reference untouched", "SUM(Table1[Amount])+A1", 1, 0, "SUM(Table1[Amount])+A2"},
		{"escaped ] in column name", "Table1[Q1']A1]+B1", 1, 0, "Table1[Q1']A1]+B2"},
		{"escaped [ in column name", "Table1[a'[b]+A1", 1, 0, "Table1[a'[b]+A2"},
		{"this-row structured reference", "Table1[[#This Row],[Qty]]*A1", 1, 0, "Table1[[#This Row],[Qty]]*A2"},
		{"defined name untouched", "TaxRate*A1", 1, 0, "TaxRate*A2"},
		{"numbers untouched", "A1*1.5+2E3", 1, 0, "A2*1.5+2E3"},
		{"off the sheet becomes #REF!", "A1+B2", -1, 0, "#REF!+B1"},
		{"last column", "XFD1+XFC1", 0, 1, "#REF!+XFD1"},
		{"no offset", "Inputs!A1", 0, 0, "Inputs!A1"},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			if got := shiftFormula(c.formula, c.dRow, c.dCol); got != c.want {
				t.Errorf("shiftFormula(%q, %d, %d) = %q, want %q", c.formula, c.dRow, c.dCol, got, c.want)
			}
		})
	}
}

// The child cells of a shared formula get the master's formula moved by their offset, read straight from the sheet XML.
func TestExtractSheetExpandsSharedFormula(t *testing.T) {
	sheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="B1"><f t="shared" ref="B1:C3" si="0">Inputs!A1+A1+$A$1</f><v>1</v></c><c r="C1"><f t="shared" si="0"/><v>2</v></c></row>
<row r="3"><c r="B3"><f t="shared" si="0"/><v>3</v></c><c r="D3"><f t="shared" si="7"/><v>4</v></c></row>
<row r="4"><c r="B4"><f t="shared" si="0"/><v>5</v></c><c r="C4"><f t="shared"/><v>6</v></c></row>
<row r="5"><c r="E5"><f t="shared" ref="A1:F9" si="">Z1</f><v>7</v></c></row>
</sheetData></worksheet>`
	zr := zipReaderFrom(t, map[string]string{"xl/worksheets/sheet1.xml": sheet})
	sh, err := extractSheet(zr, "xl/worksheets/sheet1.xml", "S", nil)
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]string{
		"B1": "Inputs!A1+A1+$A$1",
		"C1": "Inputs!B1+B1+$A$1",
		"B3": "Inputs!A3+A3+$A$1",
		"D3": "", // master si=7 is missing: keep the formula empty rather than guess
		"B4": "", // outside the master's ref B1:C3
		"C4": "", // no si: do not attach it to a master with si=""
	}
	got := map[string]string{}
	for _, c := range sh.Cells {
		got[c.Addr] = c.Formula
	}
	for addr, f := range want {
		if got[addr] != f {
			t.Errorf("%s formula = %q, want %q", addr, got[addr], f)
		}
	}
}

// The ref of a shared formula only drops children that are clearly outside it. Odd but readable refs are normalized,
// and an unreadable ref drops nothing (an earlier version collapsed it to the master cell and lost every child).
func TestSharedMasterRef(t *testing.T) {
	cases := []struct {
		ref      string
		row, col int
		want     bool
	}{
		{"B1:C3", 3, 3, true},
		{"B1:C3", 4, 2, false},
		{"$B$1:$C$3", 3, 3, true},
		{"b1:c3", 3, 3, true},  // lower case
		{"C3:B1", 3, 3, true},  // reversed corners
		{"C1:B3", 1, 2, true},  // the other diagonal
		{"C1:B3", 4, 2, false}, // still bounded
		{"B1", 1, 2, true},     // single cell
		{"B1", 2, 2, false},
		{"", 9, 9, true},        // missing ref: cannot tell, keep the child
		{"garbage", 9, 9, true}, // unreadable ref: cannot tell, keep the child
	}
	for _, c := range cases {
		m := newSharedMaster("A1", 1, 2, c.ref)
		if got := m.contains(c.row, c.col); got != c.want {
			t.Errorf("ref %q contains(%d,%d) = %v, want %v", c.ref, c.row, c.col, got, c.want)
		}
	}
}
