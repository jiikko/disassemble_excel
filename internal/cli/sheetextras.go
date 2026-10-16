package cli

import (
	"archive/zip"
	"encoding/xml"
	"path/filepath"
	"strconv"
	"strings"
)

// Sheet-level parts besides cells that carry logic or intent: data validations, hyperlinks, tables and comments.
// Presentation (merged cells, conditional formatting, styles) is out of scope: this tool disassembles logic. The element
// layouts follow excelize's xmlWorksheet.go / xmlTable.go / xmlComments.go, including the Excel 2010+ (x14) data
// validations that Excel moves into <extLst> when a list refers to another sheet. These are read in the same streaming pass as the cells (extractSheet), not through excelize, which
// would load each worksheet into memory a second time.

// DataValidation is one input rule (main <dataValidation> or x14 <x14:dataValidation>).
type DataValidation struct {
	Sqref      string
	Type       string // none (any value) | whole | decimal | list | date | time | textLength | custom
	Operator   string // between (default) | notBetween | equal | ... ; meaningless for list / custom
	Formula1   string
	Formula2   string
	AllowBlank bool
}

// Hyperlink is one <hyperlink>: an external target (from the sheet rels) or a location inside the workbook.
type Hyperlink struct {
	Cell     string
	Target   string // external URL / file ("" for an in-workbook link)
	Location string // e.g. Sheet2!A1
	Display  string
}

// TableColumn is one column of a table (ListObject). CalculatedFormula is the column formula Excel fills into every
// row; excelize does not model it.
type TableColumn struct {
	Table, Sheet, Ref, Column, CalculatedFormula string
}

// Comment is one legacy comment (note). Excel also writes this fallback for threaded comments, so their text is here
// too, prefixed by Excel's "[Threaded comment]" notice.
type Comment struct {
	Sheet, Cell, Author, Text string
}

// --- elements met while streaming the worksheet XML ---

type xmlFormulaText struct {
	Text string `xml:",chardata"`
	F    string `xml:"f"` // x14 wraps the formula in <xm:f>
}

// value prefers the element's own text (main form) and falls back to <xm:f> (x14 form, whose own text is whitespace).
func (f xmlFormulaText) value() string {
	if strings.TrimSpace(f.Text) != "" {
		return f.Text
	}
	return f.F
}

// A rule written in both the main and the x14 form appears twice, as the file has it.
// The same struct decodes the main and the x14 element: encoding/xml matches on local names here, and x14 moves
// sqref from an attribute into an <xm:sqref> child.
type xmlDataValidation struct {
	Type       string         `xml:"type,attr"`
	Operator   string         `xml:"operator,attr"`
	AllowBlank string         `xml:"allowBlank,attr"`
	SqrefAttr  string         `xml:"sqref,attr"`
	SqrefEl    string         `xml:"sqref"`
	Formula1   xmlFormulaText `xml:"formula1"`
	Formula2   xmlFormulaText `xml:"formula2"`
}

func (d xmlDataValidation) toModel() DataValidation {
	typ := d.Type
	if typ == "" {
		typ = "none"
	}
	op := d.Operator
	if op == "" && (typ == "whole" || typ == "decimal" || typ == "date" || typ == "time" || typ == "textLength") {
		op = "between" // the schema default (ECMA-376 Part 1 §18.3.1.32)
	}
	sq := d.SqrefAttr
	if sq == "" {
		sq = d.SqrefEl
	}
	return DataValidation{
		Sqref: strings.TrimSpace(sq), Type: typ, Operator: op,
		Formula1: d.Formula1.value(), Formula2: d.Formula2.value(),
		AllowBlank: d.AllowBlank == "1" || d.AllowBlank == "true",
	}
}

type xmlHyperlink struct {
	Ref      string `xml:"ref,attr"`
	RID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Location string `xml:"location,attr"`
	Display  string `xml:"display,attr"`
}

// sheetParts collects the non-cell elements of one worksheet during extractSheet.
type sheetParts struct {
	DataValidations []DataValidation
	hyperlinks      []xmlHyperlink
	tableRIDs       []string
}

const (
	nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	nsX14  = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
	nsRel  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// handle decodes se if it is one of the elements sheetParts collects (others are left to the caller's token loop).
// Elements are matched by namespace too: an extension (<extLst>) may carry same-named elements of another vocabulary.
func (p *sheetParts) handle(dec *xml.Decoder, se xml.StartElement) {
	switch ns := se.Name.Space; {
	case ns == nsMain:
	case ns == nsX14 && se.Name.Local == "dataValidation":
	default:
		return
	}
	switch se.Name.Local {
	case "dataValidation":
		var d xmlDataValidation
		if dec.DecodeElement(&d, &se) == nil {
			p.DataValidations = append(p.DataValidations, d.toModel())
		}
	case "hyperlink":
		var h xmlHyperlink
		if dec.DecodeElement(&h, &se) == nil {
			p.hyperlinks = append(p.hyperlinks, h)
		}
	case "tablePart":
		for _, a := range se.Attr {
			if a.Name.Space == nsRel && a.Name.Local == "id" {
				p.tableRIDs = append(p.tableRIDs, a.Value)
			}
		}
	}
}

// resolve fills in what needs the sheet's rels: hyperlink targets, tables and comments.
func (p *sheetParts) resolve(zr *zip.Reader, wsTarget, sheet string) (links []Hyperlink, cols []TableColumn, comments []Comment) {
	rels := sheetRels(zr, wsTarget)
	byID := map[string]rel{}
	for _, r := range rels {
		byID[r.ID] = r
	}
	for _, h := range p.hyperlinks {
		l := Hyperlink{Cell: h.Ref, Location: h.Location, Display: h.Display}
		if r, ok := byID[h.RID]; ok && h.RID != "" && strings.HasSuffix(r.Type, "/hyperlink") {
			l.Target = r.Target
		}
		links = append(links, l)
	}
	for _, id := range p.tableRIDs {
		if r, ok := byID[id]; ok && !r.External && strings.HasSuffix(r.Type, "/table") {
			cols = append(cols, parseTable(zr, r.Target, sheet)...)
		}
	}
	for _, r := range rels {
		if !r.External && strings.HasSuffix(r.Type, "/comments") {
			comments = append(comments, parseComments(zr, r.Target, sheet)...)
		}
	}
	return links, cols, comments
}

func parseTable(zr *zip.Reader, target, sheet string) []TableColumn {
	f := openZip(zr, target)
	if f == nil {
		return nil
	}
	defer func() { _ = f.Close() }()
	var t struct {
		Name        string `xml:"name,attr"`
		DisplayName string `xml:"displayName,attr"`
		Ref         string `xml:"ref,attr"`
		Columns     []struct {
			Name    string `xml:"name,attr"`
			Formula string `xml:"calculatedColumnFormula"`
		} `xml:"tableColumns>tableColumn"`
	}
	if xml.NewDecoder(f).Decode(&t) != nil {
		return nil
	}
	name := t.DisplayName // the name formulas use (Sales[Amount]); name is an internal id that usually matches
	if name == "" {
		name = t.Name
	}
	out := make([]TableColumn, 0, len(t.Columns))
	for _, c := range t.Columns {
		out = append(out, TableColumn{Table: name, Sheet: sheet, Ref: t.Ref, Column: unescapeOOXML(c.Name), CalculatedFormula: c.Formula})
	}
	return out
}

func parseComments(zr *zip.Reader, target, sheet string) []Comment {
	f := openZip(zr, target)
	if f == nil {
		return nil
	}
	defer func() { _ = f.Close() }()
	var c struct {
		Authors []string `xml:"authors>author"`
		List    []struct {
			Ref      string `xml:"ref,attr"`
			AuthorID string `xml:"authorId,attr"` // a string: one malformed value must not fail the whole part
			Text     isXML  `xml:"text"`
		} `xml:"commentList>comment"`
	}
	if xml.NewDecoder(f).Decode(&c) != nil {
		return nil
	}
	out := make([]Comment, 0, len(c.List))
	for _, it := range c.List {
		author := "" // a missing or unknown authorId leaves the author empty rather than guessing one
		if i, err := strconv.Atoi(it.AuthorID); err == nil && i >= 0 && i < len(c.Authors) {
			author = unescapeOOXML(c.Authors[i])
		}
		out = append(out, Comment{Sheet: sheet, Cell: it.Ref, Author: author, Text: it.Text.text()})
	}
	return out
}

// --- across sheets ---

// extras gathers the sheet parts of every extracted sheet and writes one TSV per kind (only kinds that occur).
type extras struct {
	validations, links, tableCols, comments [][]string
}

func (e *extras) add(zr *zip.Reader, ws wbSheet, sh *Sheet) {
	p := &sh.Parts
	for _, d := range p.DataValidations {
		e.validations = append(e.validations, []string{ws.Name, d.Sqref, d.Type, d.Operator, d.Formula1, d.Formula2, boolField(d.AllowBlank)})
	}
	links, cols, comments := p.resolve(zr, ws.Target, ws.Name)
	for _, l := range links {
		e.links = append(e.links, []string{ws.Name, l.Cell, l.Target, l.Location, l.Display})
	}
	for _, c := range cols {
		e.tableCols = append(e.tableCols, []string{c.Table, c.Sheet, c.Ref, c.Column, c.CalculatedFormula})
	}
	for _, c := range comments {
		e.comments = append(e.comments, []string{c.Sheet, c.Cell, c.Author, c.Text})
	}
}

func boolField(b bool) string {
	if b {
		return "1"
	}
	return ""
}

// write writes the TSV files and returns how many rows each got (for manifest.json / README.md).
func (e *extras) write(outDir string) (Counts, error) {
	var c Counts
	files := []struct {
		name   string
		header []string
		rows   [][]string
		count  *int
	}{
		{"data_validations.tsv", []string{"sheet", "sqref", "type", "operator", "formula1", "formula2", "allow_blank"}, e.validations, &c.DataValidations},
		{"hyperlinks.tsv", []string{"sheet", "cell", "target", "location", "display"}, e.links, &c.Hyperlinks},
		{"tables.tsv", []string{"table", "sheet", "ref", "column", "calculated_formula"}, e.tableCols, &c.TableColumns},
		{"comments.tsv", []string{"sheet", "cell", "author", "text"}, e.comments, &c.Comments},
	}
	for _, f := range files {
		*f.count = len(f.rows)
		if len(f.rows) == 0 {
			continue
		}
		if err := writeTSV(filepath.Join(outDir, f.name), f.header, f.rows); err != nil {
			return c, err
		}
	}
	return c, nil
}
