package cli

import (
	"reflect"
	"strings"
	"testing"
)

func TestUnescapeOOXML(t *testing.T) {
	cases := []struct{ in, want string }{
		{"plain", "plain"},
		{"CR_x000D_LF", "CR\rLF"},
		{"_x0009_tab", "\ttab"},
		{"lower _x000d_ hex", "lower \r hex"},
		{"literal _x005F_x0041_", "literal _x0041_"}, // _x005F_ protects a literal escape
		{"_x005F_", "_"},
		{"pair _xD83D__xDE00_", "pair 😀"}, // UTF-16 surrogate pair becomes one rune
		{"lone _xD83D_ x", "lone � x"},    // a lone surrogate cannot be encoded: U+FFFD, as Go does
		{"not an escape _x12G4_ _x123_ _x1234", "not an escape _x12G4_ _x123_ _x1234"},
		{"日本語_x000A_改行", "日本語\n改行"},
	}
	for _, c := range cases {
		if got := unescapeOOXML(c.in); got != c.want {
			t.Errorf("unescapeOOXML(%q) = %q, want %q", c.in, got, c.want)
		}
	}
}

const workbookXML = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
 <sheet name="Main" sheetId="1" r:id="rId1"/>
 <sheet name="Workbook" sheetId="2" state="hidden" r:id="rId2"/>
 <sheet name="Logic" sheetId="3" state="veryHidden" r:id="rId3"/>
</sheets>
<definedNames>
 <definedName name="Rate">0.1</definedName>
 <definedName name="Local" localSheetId="1">Workbook!$A$1</definedName>
 <definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Logic!$A$1:$B$9</definedName>
 <definedName name="Bad" localSheetId="9">#REF!</definedName>
</definedNames>
</workbook>`

const workbookRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`

func TestParseWorkbook(t *testing.T) {
	zr := zipReaderFrom(t, map[string]string{"xl/workbook.xml": workbookXML, "xl/_rels/workbook.xml.rels": workbookRels})
	wb, err := parseWorkbook(zr)
	if err != nil {
		t.Fatal(err)
	}
	wantSheets := []wbSheet{
		{Name: "Main", Target: "xl/worksheets/sheet1.xml", State: "visible"},
		{Name: "Workbook", Target: "xl/worksheets/sheet2.xml", State: "hidden"},
		{Name: "Logic", Target: "xl/worksheets/sheet3.xml", State: "veryHidden"},
	}
	if !reflect.DeepEqual(wb.Sheets, wantSheets) {
		t.Errorf("sheets = %+v\nwant %+v", wb.Sheets, wantSheets)
	}
	wantNames := []DefinedName{
		{Name: "Rate", Scope: "workbook", RefersTo: "0.1"},
		// a sheet literally named "Workbook" must not be confused with the workbook scope (excelize spells both "Workbook")
		{Name: "Local", Scope: "Workbook", RefersTo: "Workbook!$A$1"},
		{Name: "_xlnm._FilterDatabase", Scope: "Logic", RefersTo: "Logic!$A$1:$B$9", Hidden: true},
		{Name: "Bad", Scope: "[localSheetId=9]", RefersTo: "#REF!"}, // out of range: reported, not guessed as workbook
	}
	if !reflect.DeepEqual(wb.DefinedNames, wantNames) {
		t.Errorf("defined names = %+v\nwant %+v", wb.DefinedNames, wantNames)
	}
}

// One worksheet with every collected part; x14 data validations live in <extLst> with sqref as a child element.
const extrasSheetXML = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a_x000D_b</t></is></c></row></sheetData>
<dataValidations count="2">
 <dataValidation type="whole" allowBlank="1" sqref="A2:A10 C2"><formula1>1</formula1><formula2>10</formula2></dataValidation>
 <dataValidation type="custom" sqref="B2"><formula1>LEN(B2)&lt;=5<f>junk</f></formula1></dataValidation>
</dataValidations>
<hyperlinks>
 <hyperlink ref="D1" r:id="rId1" display="spec"/>
 <hyperlink ref="D2" location="'Other Sheet'!A1"/>
 <hyperlink ref="D3" r:id="rId3"/>
</hyperlinks>
<tableParts count="1"><tablePart r:id="rId2" xmlns:q="urn:other" q:id="rId2"/></tableParts>
<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
 <x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>E2:E9</xm:sqref></x14:dataValidation>
 </x14:dataValidations>
</ext>
<ext uri="{urn-other}" xmlns:foo="urn:other"><foo:hyperlink ref="Q1" location="Hidden!A1"/><foo:dataValidation type="whole" sqref="Q2"/></ext></extLst>
</worksheet>`

const extrasSheetRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a?b=1" TargetMode="External"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table7.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments3.xml"/>
</Relationships>`

const extrasTableXML = `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="Table7" displayName="Orders" ref="A1:B4">
<tableColumns count="2"><tableColumn id="1" name="Qty"/><tableColumn id="2" name="Line_x000A_Total"><calculatedColumnFormula>Orders[[#This Row],[Qty]]*2</calculatedColumnFormula></tableColumn></tableColumns></table>`

const extrasCommentsXML = `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>al_x0069_ce</author></authors>
<commentList>
 <comment ref="A1" authorId="0"><text><r><t>first </t></r><r><t>line_x000D_</t></r></text></comment>
 <comment ref="B1" authorId="5"><text><t>unknown author</t></text></comment>
 <comment ref="C1" authorId="abc"><text><t>bad id</t></text></comment>
 <comment ref="D1"><text><t>head </t><r><t>run</t></r><rPh sb="0" eb="1"><t>ふりがな</t></rPh></text></comment>
</commentList></comments>`

func TestSheetExtras(t *testing.T) {
	zr := zipReaderFrom(t, map[string]string{
		"xl/worksheets/sheet1.xml":            extrasSheetXML,
		"xl/worksheets/_rels/sheet1.xml.rels": extrasSheetRels,
		"xl/tables/table7.xml":                extrasTableXML,
		"xl/comments3.xml":                    extrasCommentsXML,
	})
	sh, err := extractSheet(zr, "xl/worksheets/sheet1.xml", "S", nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := sh.Cells[0].Value; got != "a\rb" {
		t.Errorf("inline string = %q, want %q", got, "a\rb")
	}
	wantDV := []DataValidation{
		{Sqref: "A2:A10 C2", Type: "whole", Operator: "between", Formula1: "1", Formula2: "10", AllowBlank: true},
		{Sqref: "B2", Type: "custom", Formula1: "LEN(B2)<=5"},
		{Sqref: "E2:E9", Type: "list", Formula1: "Lists!$A$1:$A$3", AllowBlank: true},
	}
	if !reflect.DeepEqual(sh.Parts.DataValidations, wantDV) {
		t.Errorf("data validations = %+v\nwant %+v", sh.Parts.DataValidations, wantDV)
	}

	links, cols, comments := sh.Parts.resolve(zr, "xl/worksheets/sheet1.xml", "S")
	wantLinks := []Hyperlink{
		{Cell: "D1", Target: "https://example.com/a?b=1", Display: "spec"},
		{Cell: "D2", Location: "'Other Sheet'!A1"},
		{Cell: "D3"}, // r:id points at a comments rel, not a hyperlink: no target rather than a wrong one
	}
	if !reflect.DeepEqual(links, wantLinks) {
		t.Errorf("hyperlinks = %+v\nwant %+v", links, wantLinks)
	}
	wantCols := []TableColumn{
		{Table: "Orders", Sheet: "S", Ref: "A1:B4", Column: "Qty"},
		{Table: "Orders", Sheet: "S", Ref: "A1:B4", Column: "Line\nTotal", CalculatedFormula: "Orders[[#This Row],[Qty]]*2"},
	}
	if !reflect.DeepEqual(cols, wantCols) {
		t.Errorf("table columns = %+v\nwant %+v", cols, wantCols)
	}
	wantComments := []Comment{
		{Sheet: "S", Cell: "A1", Author: "alice", Text: "first line\r"},
		{Sheet: "S", Cell: "B1", Text: "unknown author"}, // authorId out of range: no author rather than a crash
		{Sheet: "S", Cell: "C1", Text: "bad id"},         // a malformed authorId must not drop the other comments
		{Sheet: "S", Cell: "D1", Text: "head run"},       // <t> and runs are joined; furigana (rPh) is not text
	}
	if !reflect.DeepEqual(comments, wantComments) {
		t.Errorf("comments = %+v\nwant %+v", comments, wantComments)
	}
}

// Parts that are missing or broken (a file from outside may be anything) are skipped without failing the sheet.
func TestSheetExtrasBrokenParts(t *testing.T) {
	zr := zipReaderFrom(t, map[string]string{
		"xl/worksheets/sheet1.xml":            extrasSheetXML,
		"xl/worksheets/_rels/sheet1.xml.rels": extrasSheetRels,
		"xl/tables/table7.xml":                "<table", // truncated
		// comments3.xml is missing
	})
	sh, err := extractSheet(zr, "xl/worksheets/sheet1.xml", "S", nil)
	if err != nil {
		t.Fatal(err)
	}
	links, cols, comments := sh.Parts.resolve(zr, "xl/worksheets/sheet1.xml", "S")
	if len(links) != 3 || len(cols) != 0 || len(comments) != 0 {
		t.Errorf("got %d links, %d table columns, %d comments; want 3, 0, 0", len(links), len(cols), len(comments))
	}
}

// A zip that is not a readable workbook is an error, not an empty success (it would read as "every sheet deleted").
func TestParseWorkbookErrors(t *testing.T) {
	for name, files := range map[string]map[string]string{
		"missing":        {"hello.txt": "hi"},
		"malformed":      {"xl/workbook.xml": strings.Replace(workbookXML, "</sheets>", "</sheetz>", 1), "xl/_rels/workbook.xml.rels": workbookRels},
		"strict":         {"xl/workbook.xml": strings.Replace(workbookXML, "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main", 1), "xl/_rels/workbook.xml.rels": workbookRels},
		"not a workbook": {"xl/workbook.xml": `<foo xmlns="urn:x"><bar/></foo>`},
		"dangling sheet": {"xl/workbook.xml": workbookXML, "xl/_rels/workbook.xml.rels": strings.Replace(workbookRels, `Id="rId3"`, `Id="rId9"`, 1)},
	} {
		if _, err := parseWorkbook(zipReaderFrom(t, files)); err == nil {
			t.Errorf("%s: parseWorkbook returned no error", name)
		}
	}
}

func TestTSVEscapesControlCharacters(t *testing.T) {
	in := "a\x00b\tc\x1fd\x7fe\\x00"
	want := `a\x00b\tc\x1Fd\x7Fe\\x00`
	if got := tsvEsc(in); got != want {
		t.Errorf("tsvEsc(%q) = %q, want %q", in, got, want)
	}
	if got := unescapeOOXML("_x0000_"); got != "\x00" {
		t.Fatalf("precondition: _x0000_ decodes to NUL, got %q", got)
	}
}

// A sheet part that is missing or breaks off midway, and a truncated sharedStrings.xml, are errors: a partial result
// would read as deleted cells in a diff.
func TestTruncatedPartsAreErrors(t *testing.T) {
	cut := extrasSheetXML[:strings.Index(extrasSheetXML, "</sheetData>")]
	zr := zipReaderFrom(t, map[string]string{"xl/worksheets/sheet1.xml": cut})
	if _, err := extractSheet(zr, "xl/worksheets/sheet1.xml", "S", nil); err == nil {
		t.Error("truncated sheet: no error")
	}
	if _, err := extractSheet(zr, "xl/worksheets/sheet9.xml", "S", nil); err == nil {
		t.Error("missing sheet part: no error")
	}
	for _, sst := range []string{
		`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>a</t></si><si><t>b`, // inside an <si>
		`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>a</t></si><s`,       // between elements
	} {
		if _, err := parseSharedStrings(zipReaderFrom(t, map[string]string{"xl/sharedStrings.xml": sst})); err == nil {
			t.Errorf("truncated sharedStrings %q: no error", sst[len(sst)-12:])
		}
	}
	if got, err := parseSharedStrings(zipReaderFrom(t, map[string]string{"x": "y"})); err != nil || got != nil {
		t.Errorf("no sharedStrings part: got %v, %v; want nil, nil", got, err)
	}
}
