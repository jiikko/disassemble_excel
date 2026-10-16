// Command disassemble_excel takes one .xlsx/.xlsm file and writes its sheets
// (one cell per line, formula + cached value), defined names and VBA macro
// source into a directory, so the result can be diffed and read without Excel.
//
// It is specialized for Excel disassembly only and has no external-tool
// dependency: VBA is decompressed in pure Go (see internal/ovba).
package main

import "github.com/jiikko/disassemble_excel/internal/cli"

func main() { cli.Main() }
