package cli

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"path"
	"strconv"
	"strings"
)

// wbSheet is one entry of <sheets> in xl/workbook.xml, in workbook order.
type wbSheet struct {
	Name   string
	Target string // e.g. "xl/worksheets/sheet3.xml"
	State  string // visible | hidden | veryHidden
}

// DefinedName is one workbook/sheet-scoped name (named range, LAMBDA arg, ...).
type DefinedName struct {
	Name     string
	Scope    string // "workbook" or a sheet name
	RefersTo string
	Hidden   bool
}

type workbook struct {
	Sheets       []wbSheet
	DefinedNames []DefinedName
}

// parseWorkbook reads xl/workbook.xml (+ its rels) directly. A missing or malformed workbook.xml is an error: without
// it nothing tells which parts are sheets, and a partly decoded one would silently drop sheets or names. excelize's GetDefinedName does not report hidden names and
// spells the workbook scope "Workbook", which collides with a sheet named Workbook; GetSheetVisible cannot tell
// hidden from veryHidden. Reading the XML ourselves keeps both.
func parseWorkbook(zr *zip.Reader) (workbook, error) {
	type sheetEl struct {
		Name  string `xml:"name,attr"`
		State string `xml:"state,attr"`
		RID   string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	}
	type nameEl struct {
		Name         string `xml:"name,attr"`
		LocalSheetID string `xml:"localSheetId,attr"`
		Hidden       string `xml:"hidden,attr"`
		Text         string `xml:",chardata"`
	}
	var wb struct {
		XMLName xml.Name
		Sheets  []sheetEl `xml:"sheets>sheet"`
		Names   []nameEl  `xml:"definedNames>definedName"`
	}
	f := openZip(zr, "xl/workbook.xml")
	if f == nil {
		return workbook{}, errors.New("xl/workbook.xml が無い (Excel のワークブックではない)")
	}
	err := xml.NewDecoder(f).Decode(&wb)
	_ = f.Close()
	if err != nil {
		return workbook{}, fmt.Errorf("xl/workbook.xml を読めない: %w", err)
	}
	if wb.XMLName.Space == nsStrictMain {
		return workbook{}, errors.New("この形式 (Strict Open XML) のワークブックは未対応 (Excel で「Excel ブック (.xlsx)」として保存し直す)")
	}
	if wb.XMLName.Local != "workbook" {
		return workbook{}, fmt.Errorf("xl/workbook.xml のルートが <%s> で、ワークブックではない", wb.XMLName.Local)
	}
	ridTarget := map[string]string{}
	for _, r := range readRels(zr, "xl/_rels/workbook.xml.rels", "xl") {
		ridTarget[r.ID] = r.Target
	}

	var out workbook
	for _, s := range wb.Sheets {
		state := s.State
		if state == "" {
			state = "visible"
		}
		t := ridTarget[s.RID]
		if t == "" {
			// skipping it would make the sheet look deleted in a diff
			return workbook{}, fmt.Errorf("シート %q の実体 (r:id=%q) が xl/_rels/workbook.xml.rels に無い", s.Name, s.RID)
		}
		out.Sheets = append(out.Sheets, wbSheet{Name: s.Name, Target: t, State: state})
	}
	for _, n := range wb.Names {
		// localSheetId indexes <sheets> in document order (ECMA-376 Part 1 §18.2.5), including sheets without a target.
		// One that points at no sheet is reported as is ([] cannot occur in a sheet name) rather than guessed as workbook.
		// Known gap: a sheet literally named "workbook" is indistinguishable from the workbook scope (as in earlier
		// versions); changing the spelling of the workbook scope would change every existing defined_names.tsv.
		scope := "workbook"
		if n.LocalSheetID != "" {
			scope = "[localSheetId=" + n.LocalSheetID + "]"
			if i, err := strconv.Atoi(n.LocalSheetID); err == nil && i >= 0 && i < len(wb.Sheets) {
				scope = wb.Sheets[i].Name
			}
		}
		out.DefinedNames = append(out.DefinedNames, DefinedName{
			Name: n.Name, Scope: scope, RefersTo: n.Text, Hidden: n.Hidden == "1" || n.Hidden == "true",
		})
	}
	return out, nil
}

const nsStrictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main"

// rel is one Relationship of a .rels part. Target is resolved to a zip path unless External.
type rel struct {
	ID, Type, Target string
	External         bool
}

// readRels reads a .rels part; baseDir is the directory of the part the rels belong to (targets are relative to it).
func readRels(zr *zip.Reader, relsPath, baseDir string) []rel {
	f := openZip(zr, relsPath)
	if f == nil {
		return nil
	}
	defer func() { _ = f.Close() }()
	var rels struct {
		Rel []struct {
			ID         string `xml:"Id,attr"`
			Type       string `xml:"Type,attr"`
			Target     string `xml:"Target,attr"`
			TargetMode string `xml:"TargetMode,attr"`
		} `xml:"Relationship"`
	}
	if xml.NewDecoder(f).Decode(&rels) != nil {
		return nil
	}
	out := make([]rel, 0, len(rels.Rel))
	for _, r := range rels.Rel {
		ext := strings.EqualFold(r.TargetMode, "External")
		t := r.Target
		if !ext {
			t = resolveRelTarget(baseDir, t)
		}
		out = append(out, rel{ID: r.ID, Type: r.Type, Target: t, External: ext})
	}
	return out
}

// sheetRels reads the .rels of one worksheet part.
func sheetRels(zr *zip.Reader, wsTarget string) []rel {
	dir := path.Dir(wsTarget)
	return readRels(zr, dir+"/_rels/"+path.Base(wsTarget)+".rels", dir)
}
