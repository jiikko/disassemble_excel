package cli

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf16"
)

// --- A1 <-> (row, col) helpers ---

func colToNum(col string) int {
	n := 0
	for i := 0; i < len(col); i++ {
		n = n*26 + int(col[i]-'A'+1)
	}
	return n
}

func numToCol(n int) string {
	var b []byte
	for n > 0 {
		n--
		b = append([]byte{byte('A' + n%26)}, b...)
		n /= 26
	}
	return string(b)
}

func splitAddr(addr string) (row, col int) {
	i := 0
	for i < len(addr) && addr[i] >= 'A' && addr[i] <= 'Z' {
		i++
	}
	col = colToNum(addr[:i])
	row, _ = strconv.Atoi(addr[i:])
	return row, col
}

// --- shared strings ---

type siXML struct {
	T string     `xml:"t"`
	R []siRunXML `xml:"r"`
}
type siRunXML struct {
	T string `xml:"t"`
}

func (s siXML) text() string { return richText(s.T, s.R) }

// richText joins a CT_Rst: an optional plain <t> followed by rich-text runs (both may appear; excelize joins them too).
// Phonetic runs (<rPh>) are not part of the text.
func richText(t string, runs []siRunXML) string {
	var b strings.Builder
	b.WriteString(t)
	for _, r := range runs {
		b.WriteString(r.T)
	}
	return unescapeOOXML(b.String())
}

// unescapeOOXML decodes the _xHHHH_ escapes of an ST_Xstring (ECMA-376 Part 1 §22.9.2.19): characters XML 1.0 cannot
// carry (CR, control characters) are stored as _x000D_ etc., and a literal "_x" sequence is protected as _x005F_x....
// Same rule as excelize's bstrUnmarshal, except that a UTF-16 surrogate pair (_xD83D__xDE00_) becomes one character.
// Applied to shared and inline strings (CT_Rst) after joining the runs, as excelize does; the cached value of a
// formula string (t="str") is left as stored, also as excelize does. Control characters it yields are escaped by the
// writers (escapeControls), not here.
func unescapeOOXML(s string) string {
	if !strings.Contains(s, "_x") {
		return s
	}
	var b strings.Builder
	for i := 0; i < len(s); {
		if r, ok := escapedUnit(s, i); ok {
			if r == '_' { // _x005F_ escapes the underscore of a literal "_xHHHH_"
				b.WriteByte('_')
				i += 7
				continue
			}
			if utf16.IsSurrogate(r) {
				if r2, ok2 := escapedUnit(s, i+7); ok2 {
					if pair := utf16.DecodeRune(r, r2); pair != unicode.ReplacementChar {
						b.WriteRune(pair)
						i += 14
						continue
					}
				}
			}
			b.WriteRune(r)
			i += 7
			continue
		}
		b.WriteByte(s[i])
		i++
	}
	return b.String()
}

// escapedUnit reads one _xHHHH_ at s[i:].
func escapedUnit(s string, i int) (rune, bool) {
	if i+7 > len(s) || s[i] != '_' || s[i+1] != 'x' || s[i+6] != '_' {
		return 0, false
	}
	n, err := strconv.ParseUint(s[i+2:i+6], 16, 16)
	if err != nil {
		return 0, false
	}
	return rune(n), true
}

// parseSharedStrings reads xl/sharedStrings.xml. A workbook without strings has no such part; a part that breaks off
// midway is an error, because every later string index would silently resolve to "".
func parseSharedStrings(zr *zip.Reader) ([]string, error) {
	f := openZip(zr, "xl/sharedStrings.xml")
	if f == nil {
		return nil, nil
	}
	defer func() { _ = f.Close() }()
	dec := xml.NewDecoder(f)
	var out []string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, fmt.Errorf("xl/sharedStrings.xml を読めない: %w", err)
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "si" {
			var si siXML
			if err := dec.DecodeElement(&si, &se); err != nil {
				return nil, fmt.Errorf("xl/sharedStrings.xml を読めない: %w", err)
			}
			out = append(out, si.text())
		}
	}
}

// --- worksheet cell parsing ---

type cXML struct {
	R  string `xml:"r,attr"`
	T  string `xml:"t,attr"`
	F  *fXML  `xml:"f"`
	V  string `xml:"v"`
	Is *isXML `xml:"is"`
}
type fXML struct {
	T       string `xml:"t,attr"`
	SI      string `xml:"si,attr"`
	Ref     string `xml:"ref,attr"`
	Content string `xml:",chardata"`
}
type isXML struct {
	T string     `xml:"t"`
	R []siRunXML `xml:"r"`
}

func (is *isXML) text() string {
	if is == nil {
		return ""
	}
	return richText(is.T, is.R)
}

// extractSheet parses one worksheet XML and expands shared-formula children from their master cell.
// A missing or truncated sheet part is an error: emitting nothing or a partial sheet would read as deleted cells in a diff.
func extractSheet(zr *zip.Reader, target, name string, shared []string) (*Sheet, error) {
	f := openZip(zr, target)
	if f == nil {
		return nil, fmt.Errorf("シート %q の実体 %s が zip に無い", name, target)
	}
	defer func() { _ = f.Close() }()

	sh := &Sheet{Name: name}
	parts := &sh.Parts
	masters := map[string]sharedMaster{}
	dec := xml.NewDecoder(f)
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("シート %q (%s) を読めない: %w", name, target, err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "dimension":
			for _, a := range se.Attr {
				if a.Name.Local == "ref" {
					sh.Dimension = a.Value
				}
			}
		case "c":
			var c cXML
			if err := dec.DecodeElement(&c, &se); err != nil {
				return nil, fmt.Errorf("シート %q (%s) を読めない: %w", name, target, err)
			}
			if c.R == "" {
				continue
			}
			cell := buildCell(c, shared)
			if cell == nil {
				continue
			}
			if c.F != nil && c.F.T == "shared" && cell.Formula != "" {
				masters[c.F.SI] = newSharedMaster(cell.Formula, cell.Row, cell.Col, c.F.Ref)
			}
			sh.Cells = append(sh.Cells, *cell)
		default:
			parts.handle(dec, se)
		}
	}

	// Expand shared-formula children (their formula text is empty in the XML). Done after the whole sheet is read so
	// the result does not depend on the master appearing first. A child whose master is missing, or that lies outside
	// the master's ref, or has no si keeps an empty formula (a malformed file; better empty than a guess).
	for i := range sh.Cells {
		c := &sh.Cells[i]
		if m, ok := masters[c.sharedSI]; ok && c.needExpand && c.sharedSI != "" && m.contains(c.Row, c.Col) {
			c.Formula = shiftFormula(m.formula, c.Row-m.row, c.Col-m.col)
		}
	}

	sort.Slice(sh.Cells, func(a, b int) bool {
		if sh.Cells[a].Row != sh.Cells[b].Row {
			return sh.Cells[a].Row < sh.Cells[b].Row
		}
		return sh.Cells[a].Col < sh.Cells[b].Col
	})
	for _, c := range sh.Cells {
		if c.Formula != "" {
			sh.FormulaN++
		} else if c.Value != "" {
			sh.ValueN++
		}
	}
	return sh, nil
}

func buildCell(c cXML, shared []string) *Cell {
	row, col := splitAddr(c.R)
	if row == 0 {
		return nil
	}

	var value string
	switch c.T {
	case "s": // shared string index
		if idx, err := strconv.Atoi(strings.TrimSpace(c.V)); err == nil && idx >= 0 && idx < len(shared) {
			value = shared[idx]
		}
	case "inlineStr":
		value = c.Is.text()
	default: // n, b, str, e, d ...
		value = c.V
	}

	formula := ""
	needExpand := false
	sharedSI := ""
	if c.F != nil {
		if strings.TrimSpace(c.F.Content) != "" {
			formula = c.F.Content
		} else if c.F.T == "shared" {
			needExpand, sharedSI = true, c.F.SI
		}
	}

	if formula == "" && value == "" && !needExpand {
		return nil
	}
	typ := c.T
	if typ == "" {
		typ = "n"
	}
	return &Cell{Row: row, Col: col, Addr: c.R, Type: typ, Formula: formula, Value: value, needExpand: needExpand, sharedSI: sharedSI}
}

func openZip(zr *zip.Reader, name string) io.ReadCloser {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil
			}
			return rc
		}
	}
	return nil
}
