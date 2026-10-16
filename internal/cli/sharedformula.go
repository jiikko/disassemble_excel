package cli

import (
	"strconv"
	"strings"
	"unicode"
)

// Shared formulas: only the master cell (<f t="shared" ref=".." si="N">text</f>) carries the formula text; every child
// (<f t="shared" si="N"/>) means "the master's formula, with relative references moved by my offset from the master".
// excelize's GetCellFormula did this but left sheet-qualified references (Inputs!A1) unshifted, so children silently
// got the wrong formula (issue #2). shiftFormula moves references itself.

const (
	maxRow = 1048576
	maxCol = 16384 // XFD
)

// sharedMaster is the master cell of one shared-formula group (keyed by si).
type sharedMaster struct {
	formula  string
	row, col int
	// the group's range (ref attribute). A child outside it is not part of this group (a malformed file), so it is
	// left without a formula rather than given a guessed one.
	top, left, bottom, right int
}

func newSharedMaster(formula string, row, col int, ref string) sharedMaster {
	m := sharedMaster{formula: formula, row: row, col: col, top: row, left: col, bottom: row, right: col}
	from, to, isRange := strings.Cut(ref, ":")
	if !isRange {
		to = from
	}
	r1, c1 := splitAddr(strings.ReplaceAll(from, "$", ""))
	r2, c2 := splitAddr(strings.ReplaceAll(to, "$", ""))
	if r1 > 0 && c1 > 0 && r2 >= r1 && c2 >= c1 {
		m.top, m.left, m.bottom, m.right = r1, c1, r2, c2
	}
	return m
}

func (m sharedMaster) contains(row, col int) bool {
	return row >= m.top && row <= m.bottom && col >= m.left && col <= m.right
}

// shiftFormula returns formula with every relative A1 reference moved by (dRow, dCol). Absolute parts ($A$1, A$1, $A1)
// stay put. A reference moved off the sheet becomes #REF!, as Excel does.
//
// Left as is: string literals ("A1"), sheet / workbook prefixes (Inputs!, 'My Sheet'!, [1]Sheet!), function names
// (LOG10(), structured references (Table1[Col]) and defined names. A defined name can never look like a cell
// reference (Excel rejects such names), so a word that parses as a reference is one.
//
// Not handled, on the (unconfirmed) assumption that Excel never saves them in this form: lower-case
// references (a1), trim-range operators written literally (A1:.B10; Excel saves _xlfn._TRO_*), and the shape of a range
// whose one end leaves the sheet (A2:#REF!; Excel may turn the whole range into #REF!). These came from the adversarial
// review of issue #2; to raise one again, attach an Excel-saved file that shows the formula text.
func shiftFormula(formula string, dRow, dCol int) string {
	if dRow == 0 && dCol == 0 {
		return formula
	}
	rs := []rune(formula)
	var b strings.Builder
	for i := 0; i < len(rs); {
		switch r := rs[i]; {
		case r == '"': // string literal; "" is an escaped quote
			j := i + 1
			for j < len(rs) {
				if rs[j] == '"' {
					if j+1 < len(rs) && rs[j+1] == '"' {
						j += 2
						continue
					}
					j++
					break
				}
				j++
			}
			b.WriteString(string(rs[i:j]))
			i = j
		case r == '\'': // quoted sheet name; '' is an escaped quote
			j := i + 1
			for j < len(rs) {
				if rs[j] == '\'' {
					if j+1 < len(rs) && rs[j+1] == '\'' {
						j += 2
						continue
					}
					j++
					break
				}
				j++
			}
			b.WriteString(string(rs[i:j]))
			i = j
		case r == '[': // external workbook index or structured reference (may nest)
			depth, j := 0, i
			for j < len(rs) {
				if rs[j] == '\'' && depth > 0 && j+1 < len(rs) {
					// a structured-reference column name escapes [ ] # ' with a leading ' (Table1[Q1']A1])
					j += 2
					continue
				}
				if rs[j] == '[' {
					depth++
				} else if rs[j] == ']' {
					depth--
					if depth == 0 {
						j++
						break
					}
				}
				j++
			}
			b.WriteString(string(rs[i:j]))
			i = j
		case isWordRune(r):
			j := i
			for j < len(rs) && isWordRune(rs[j]) {
				j++
			}
			word := string(rs[i:j])
			next := rune(0)
			if j < len(rs) {
				next = rs[j]
			}
			switch {
			case next == '!' || next == '(' || next == '[':
				// sheet prefix, function name, table name
				b.WriteString(word)
				i = j
			case next == ':' && isRangeOfSheets(rs, j+1):
				// 3D reference Sheet1:Sheet3!A1 — both words are sheet names
				k := j + 1
				for k < len(rs) && isWordRune(rs[k]) {
					k++
				}
				b.WriteString(string(rs[i:k]))
				i = k
			case next == ':' && (isColOnly(word) || isRowOnly(word)):
				// whole-column (A:C) / whole-row (1:3) range: shift only when both ends are the same kind
				k := j + 1
				for k < len(rs) && isWordRune(rs[k]) {
					k++
				}
				other := string(rs[j+1 : k])
				switch {
				case isColOnly(word) && isColOnly(other):
					b.WriteString(shiftCol(word, dCol) + ":" + shiftCol(other, dCol))
					i = k
				case isRowOnly(word) && isRowOnly(other):
					b.WriteString(shiftRow(word, dRow) + ":" + shiftRow(other, dRow))
					i = k
				default:
					b.WriteString(shiftCell(word, dRow, dCol))
					i = j
				}
			default:
				b.WriteString(shiftCell(word, dRow, dCol))
				i = j
			}
		default:
			b.WriteRune(r)
			i++
		}
	}
	return b.String()
}

// isWordRune covers references, names, numbers and unquoted (also non-ASCII) sheet names.
func isWordRune(r rune) bool {
	return r == '$' || r == '_' || r == '.' || r == '\\' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

// isRangeOfSheets reports whether rs[from:] is "<word>!" (the second half of Sheet1:Sheet3!).
func isRangeOfSheets(rs []rune, from int) bool {
	k := from
	for k < len(rs) && isWordRune(rs[k]) {
		k++
	}
	return k > from && k < len(rs) && rs[k] == '!'
}

// parseRef splits "$AB$12" into its parts. ok is false unless the word is exactly one A1 reference on the sheet.
func parseRef(word string) (colAbs bool, col int, rowAbs bool, row int, ok bool) {
	s := word
	if strings.HasPrefix(s, "$") {
		colAbs, s = true, s[1:]
	}
	i := 0
	for i < len(s) && i < 3 && isUpperASCII(s[i]) {
		i++
	}
	if i == 0 || i == len(s) {
		return false, 0, false, 0, false
	}
	col = colToNum(s[:i])
	s = s[i:]
	if strings.HasPrefix(s, "$") {
		rowAbs, s = true, s[1:]
	}
	if s == "" || s[0] == '0' || !allDigits(s) {
		return false, 0, false, 0, false
	}
	row, err := strconv.Atoi(s)
	if err != nil || row > maxRow || col > maxCol {
		return false, 0, false, 0, false
	}
	return colAbs, col, rowAbs, row, true
}

func shiftCell(word string, dRow, dCol int) string {
	colAbs, col, rowAbs, row, ok := parseRef(word)
	if !ok {
		return word
	}
	if !colAbs {
		col += dCol
	}
	if !rowAbs {
		row += dRow
	}
	if col < 1 || col > maxCol || row < 1 || row > maxRow {
		return "#REF!"
	}
	return dollar(colAbs) + numToCol(col) + dollar(rowAbs) + strconv.Itoa(row)
}

func isColOnly(word string) bool {
	s := strings.TrimPrefix(word, "$")
	if s == "" || len(s) > 3 {
		return false
	}
	for i := 0; i < len(s); i++ {
		if !isUpperASCII(s[i]) {
			return false
		}
	}
	return colToNum(s) <= maxCol
}

func isRowOnly(word string) bool {
	s := strings.TrimPrefix(word, "$")
	if s == "" || s[0] == '0' || !allDigits(s) {
		return false
	}
	n, err := strconv.Atoi(s)
	return err == nil && n <= maxRow
}

func shiftCol(word string, d int) string {
	if strings.HasPrefix(word, "$") {
		return word
	}
	c := colToNum(word) + d
	if c < 1 || c > maxCol {
		return "#REF!"
	}
	return numToCol(c)
}

func shiftRow(word string, d int) string {
	if strings.HasPrefix(word, "$") {
		return word
	}
	n, _ := strconv.Atoi(word)
	n += d
	if n < 1 || n > maxRow {
		return "#REF!"
	}
	return strconv.Itoa(n)
}

func dollar(abs bool) string {
	if abs {
		return "$"
	}
	return ""
}

func isUpperASCII(c byte) bool { return c >= 'A' && c <= 'Z' }

func allDigits(s string) bool {
	for i := 0; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return s != ""
}
