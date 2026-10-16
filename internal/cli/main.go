// Package cli implements the disassemble_excel command. The repository root
// main.go only calls Main.
package cli

import (
	"archive/zip"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

const toolVersion = "0.1.0"

// Main parses os.Args, runs the disassembly and exits the process on failure.
func Main() {
	out := flag.String("o", "", "output directory (default: <file name without extension>/)")
	noVBA := flag.Bool("no-vba", false, "skip VBA macro extraction")
	noGrid := flag.Bool("no-grid", false, "skip the grid values.csv (cells.tsv is always written)")
	maxCells := flag.Int("max-cells", 100000, "sheets with more cells than this also get a .slim.tsv (0 disables slim)")
	sheetsCSV := flag.String("sheets", "", "comma-separated sheet names to extract (default: all)")
	force := flag.Bool("f", false, "overwrite the output directory if it already exists")
	flag.Usage = func() { usage(os.Stderr) }

	// Go's flag package stops parsing at the first non-flag argument, so accept
	// both "<file> -o out" and "-o out <file>" by pulling a leading file
	// argument out before parsing the rest.
	rawArgs := os.Args[1:]
	// ヘルプは求められて出す出力なので stdout へ出して rc=0 で終える (使い方の誤りのときだけ stderr・rc=2)。
	// `cmd --help | less` や、stdout だけを読む呼び出し元 (エージェント・スクリプト) からも読めるようにする
	for _, a := range rawArgs {
		if a == "-h" || a == "-help" || a == "--help" {
			usage(os.Stdout)
			os.Exit(0)
		}
	}
	var src string
	if len(rawArgs) > 0 && !strings.HasPrefix(rawArgs[0], "-") {
		src, rawArgs = rawArgs[0], rawArgs[1:]
	}
	_ = flag.CommandLine.Parse(rawArgs) // flag.ExitOnError なのでエラーは返らず exit する
	if src == "" {
		src = flag.Arg(0)
	}
	if src == "" {
		usage(os.Stderr)
		os.Exit(2)
	}

	if err := run(src, *out, *sheetsCSV, *maxCells, *noVBA, *noGrid, *force); err != nil {
		fmt.Fprintln(os.Stderr, "disassemble_excel:", err)
		os.Exit(1)
	}
}

func usage(w io.Writer) {
	_, _ = io.WriteString(w, `disassemble_excel `+toolVersion+` — Excel のシート・数式・VBA を diff 可能なテキストに分解する

.xlsx/.xlsm ワークブック 1 ファイルを、行指向のプレーンテキスト群に書き出します:
全セルの数式とキャッシュ値の両方、定義名、VBA マクロのソースコード、
図形・ボタンに割り当てられたマクロ。数式は抽出のみで評価（再計算）はしません。
共有数式は展開され、各セルが実際に計算している数式が見えます。
.xlsx/.xlsm の解体は完全オフラインで動作し、入力ファイルは変更しません。

.xlsb（バイナリ形式）は直接解体できませんが、LibreOffice (soffice) が
PATH か /Applications/LibreOffice.app に見つかれば、自動で .xlsm に変換してから
解体します（この場合のみ外部プロセスを起動します。無ければ導入方法を表示して rc=1）。

使い方:
  disassemble_excel <file.xlsx|.xlsm|.xlsb> [options]
  disassemble_excel -h | --help        このヘルプを stdout に出す
  （フラグはファイル引数の前後どちらに置いてもよい。1 回の起動で 1 ファイル）

オプション:
  -o DIR         出力先ディレクトリ（既定: カレントの <ファイル名から拡張子を除いたもの>/）
  -sheets a,b,c  指定したシートだけを抽出（既定: 全シート。名前は完全一致）
  -max-cells N   セル数が N を超えるシートには .slim.tsv も出力（既定 100000、0 で無効）
  -no-grid       グリッド形式の values.csv を出力しない（巨大シートで速く・小さくしたいとき）
  -no-vba        VBA マクロを抽出しない
  -f             出力先ディレクトリが空でなくても書き込む。上書きするだけで既存ファイルは
                 消さないので、前回の出力（消えたシート・モジュールのファイル）が残る。
                 作り直すときは出力先を消してから実行する

出力レイアウト:
  <out>/
    README.md                来歴: 元ファイルの絶対パス・SHA-256 と、各ファイルの意味
    manifest.json            元ファイルの SHA-256、シートごとの状態 (visible / hidden /
                             veryHidden)・寸法・セル数・数式数、VBA モジュール一覧、
                             マクロ割り当てオブジェクト一覧、下の TSV の件数 (counts)
    defined_names.tsv        name  scope  refers_to  hidden（名前付き範囲、LAMBDA 引数。
                             hidden は隠し名なら 1）
    sheets/<name>.cells.tsv  cell  type  formula  value — 1 セル 1 行、(行, 列) でソート、
                             空セルは省略。formula は先頭に = が付く（数式の無いセルは空）
    sheets/<name>.values.csv キャッシュ値のグリッド（目視確認用）
    sheets/<name>.slim.tsv   数式セルのみの圧縮ビュー（-max-cells 超のシートにだけ出力）
    vba/<module>.bas|.cls    コードを持つ VBA モジュールごとに 1 ファイル
    vba_index.tsv            module  proc  kind  module_line — Sub/Function/Property の索引
                             (module_line は vba/<module> の中の 1 始まりの行番号)
    objects.tsv              sheet  object  kind  anchor  macro — 画像・図形・ボタンに
                             割り当てられた（クリックで実行される）マクロ。画像自体は復元しない
    data_validations.tsv     sheet  sqref  type  operator  formula1  formula2  allow_blank
                             — 入力規則（入力できる値・リストの出どころ）
    tables.tsv               table  sheet  ref  column  calculated_formula — テーブルの
                             範囲と列。構造化参照 Sales[Amount] を範囲に戻すのに使う
    comments.tsv             sheet  cell  author  text — セルのコメント（メモ）
    hyperlinks.tsv           sheet  cell  target  location  display
  （objects.tsv 以下の TSV は該当が 1 件以上あるときだけ出る。件数は manifest.json の counts）
  （data_validations / tables の数式は XML のままで、先頭に = は付かない）

値の読み方:
  - type は Excel の生のセル型: n 数値 / s 共有文字列 / inlineStr インライン文字列 /
    str 数式が返した文字列 / b 真偽値 (1/0) / e エラー (#DIV/0! 等)
  - value は Excel がファイルに最後に保存したキャッシュ値で、再計算結果ではない。
    表示書式は適用しない: 日付は通し番号（例 46054）、パーセントは小数のまま出る
  - TSV の値の中の \ タブ 改行 CR は \\ \t \n \r にエスケープされ、1 セルは必ず 1 行
  - 出力は UTF-8。VBA のソースはプロジェクトのコードページ（日本語なら Shift_JIS）から変換済み
  - 共有文字列・インライン文字列の _x000D_ のようなエスケープ（XML に書けない文字の表現）は
    元の文字に戻してある。タブ・改行・CR 以外の制御文字は \x00 の形で書く（grep がバイナリ扱いしないように）
  - 空の VBA モジュール（シートクラスのスタブ等）はファイル化されず、
    manifest.json に "empty": true として記録される

エージェント・スクリプトから使うとき:
  - stdout には何も出さない。進捗（シートごとのセル数など）と結果の要約は stderr に、
    エラーは "disassemble_excel: <理由>" の 1 行で stderr に出る
  - 成果物はすべて出力先のファイル。まず manifest.json（全体の地図）を読み、
    数式は grep -n '=SUM' <out>/sheets/*.cells.tsv、マクロは vba_index.tsv で
    手続きの位置を引いてから vba/<module> を読む、の順が速い
  - 特定セルは grep -P '^B12\t' <out>/sheets/<name>.cells.tsv で引ける
  - 巨大シートの cells.tsv は数十 MB になりうる。全体を読まず slim.tsv や grep で絞る

抽出しないもの（見た目に属するもの。このツールはロジックを解体する）:
  書式（色・罫線・フォント・表示形式）、条件付き書式、結合セル、グラフ、ピボットテーブル、
  スライサー、画像・図形の見た目、非表示の行・列、ウィンドウ枠の固定、シートの保護、文書プロパティ。
  グループ図形の内側の子に個別に割り当てたマクロも拾わない（グループ自体の割り当ては拾う）。
  ワークブックはメモリに読み込む（~70 MB のマクロ付きで約 1 GB）。

実行例:
  # すべてを ./report/ に抽出する
  disassemble_excel report.xlsm

  # 出力先を指定し、2 シートだけ抽出する
  disassemble_excel report.xlsm -o out/report -sheets Summary,Data

  # 複数ファイルを 1 ファイル 1 プロセスで処理する
  for f in *.xlsm; do disassemble_excel "$f" -o "out/${f%.xlsm}" -f; done

  # 同じワークブックのバージョン間・変種間を比較する
  diff -ru out/report_v1/sheets out/report_v2/sheets

終了ステータス:
  0 成功（-h / --help も 0）, 1 抽出エラー（出力先が空でない・読めないファイル・
  xl/workbook.xml の無い zip 等）, 2 使い方の誤り
`)
}

func run(src, outDir, sheetsCSV string, maxCells int, noVBA, noGrid, force bool) error {
	start := time.Now()

	// .xlsb (Excel Binary Workbook) is a ZIP container of BIFF12 *binary* records,
	// not the OOXML (ZIP+XML) parts this tool reads, so excelize and the XML
	// parsers below cannot open it. If LibreOffice (soffice) is available, convert
	// it to .xlsm first (xlsm, not xlsx, to keep VBA). This is the ONLY place the
	// tool ever spawns a subprocess; the .xlsx/.xlsm path stays fully offline.
	origSrc := src
	if strings.EqualFold(filepath.Ext(src), ".xlsb") {
		converted, cleanup, err := convertXLSB(src)
		if err != nil {
			return err
		}
		defer cleanup()
		src = converted
	}

	if outDir == "" {
		base := filepath.Base(origSrc)
		outDir = strings.TrimSuffix(base, filepath.Ext(base))
	}
	if err := prepareOutDir(outDir, force); err != nil {
		return err
	}

	zr, err := zip.OpenReader(src)
	if err != nil {
		if isOLE2(src) {
			return errors.New("zip ではなく OLE2 形式のファイル: パスワードで暗号化されたブックか、旧形式の .xls。" +
				"Excel でパスワードを外すか、.xlsx / .xlsm で保存し直す")
		}
		return fmt.Errorf("open zip: %w", err)
	}
	defer func() { _ = zr.Close() }()

	var only map[string]bool
	if sheetsCSV != "" {
		only = map[string]bool{}
		for _, s := range strings.Split(sheetsCSV, ",") {
			only[strings.TrimSpace(s)] = true
		}
	}

	shared, err := parseSharedStrings(&zr.Reader)
	if err != nil {
		return err
	}
	wb, err := parseWorkbook(&zr.Reader)
	if err != nil {
		return err
	}
	wbSheets := wb.Sheets
	var ex extras

	sheetsDir := filepath.Join(outDir, "sheets")
	if err := os.MkdirAll(sheetsDir, 0o755); err != nil {
		return err
	}

	var sheetManifests []SheetManifest
	for _, ws := range wbSheets {
		if only != nil && !only[ws.Name] {
			continue
		}
		sh, err := extractSheet(&zr.Reader, ws.Target, ws.Name, shared)
		if err != nil {
			return err
		}
		ex.add(&zr.Reader, ws, sh)
		fileBase := filepath.Join(sheetsDir, sanitizeFilename(ws.Name))
		if err := writeCellsTSV(fileBase+".cells.tsv", sh); err != nil {
			return err
		}
		if !noGrid {
			if err := writeValuesCSV(fileBase+".values.csv", sh); err != nil {
				return err
			}
		}
		slim := maxCells > 0 && len(sh.Cells) > maxCells
		if slim {
			if err := writeSlimTSV(fileBase+".slim.tsv", sh); err != nil {
				return err
			}
		}
		sheetManifests = append(sheetManifests, SheetManifest{
			Name: ws.Name, Dimension: sh.Dimension, Cells: len(sh.Cells), Formulas: sh.FormulaN, Slim: slim, State: ws.State,
		})
		fmt.Fprintf(os.Stderr, "  sheet %-22s cells=%-7d formulas=%-6d%s\n",
			ws.Name, len(sh.Cells), sh.FormulaN, slimNote(slim))
	}

	if err := writeDefinedNames(filepath.Join(outDir, "defined_names.tsv"), wb.DefinedNames); err != nil {
		return err
	}
	counts, err := ex.write(outDir)
	if err != nil {
		return err
	}

	var modManifests []ModManifest
	if !noVBA {
		mods, err := extractVBAFromZip(&zr.Reader)
		if err != nil {
			fmt.Fprintf(os.Stderr, "  vba: %v (skipped)\n", err)
		} else if len(mods) > 0 {
			vbaDir := filepath.Join(outDir, "vba")
			if err := os.MkdirAll(vbaDir, 0o755); err != nil {
				return err
			}
			for _, m := range mods {
				mm := ModManifest{Name: m.Name, Type: m.Type, Empty: m.Empty, Procs: len(m.Procs), Lines: lineCount(m.Source)}
				if !m.Empty {
					if err := writeModule(vbaDir, m); err != nil {
						return err
					}
					mm.File = "vba/" + sanitizeFilename(m.Name) + m.Ext
				}
				modManifests = append(modManifests, mm)
			}
			if err := writeVBAIndex(filepath.Join(outDir, "vba_index.tsv"), mods); err != nil {
				return err
			}
			fmt.Fprintf(os.Stderr, "  vba: %d modules\n", len(mods))
		}
	}

	objects := extractObjects(&zr.Reader, wbSheets, only)
	if len(objects) > 0 {
		if err := writeObjectsTSV(filepath.Join(outDir, "objects.tsv"), objects); err != nil {
			return err
		}
		fmt.Fprintf(os.Stderr, "  objects: %d with assigned macro\n", len(objects))
	}

	man := Manifest{
		Source:      filepath.Base(origSrc),
		SHA256:      fileSHA256(origSrc),
		ToolVersion: toolVersion,
		Sheets:      sheetManifests,
		VBAModules:  modManifests,
		Objects:     objects,
		Counts:      counts,
	}
	if err := writeManifest(filepath.Join(outDir, "manifest.json"), man); err != nil {
		return err
	}

	srcAbs, err := filepath.Abs(origSrc)
	if err != nil {
		srcAbs = origSrc
	}
	if err := writeReadme(filepath.Join(outDir, "README.md"), man, srcAbs, noGrid); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "done: %s (%d sheets, %d VBA modules) in %s\n",
		outDir, len(sheetManifests), len(modManifests), time.Since(start).Round(time.Millisecond))
	return nil
}

// convertXLSB converts an .xlsb workbook to .xlsm via LibreOffice (soffice) so
// the OOXML disassembly path can read it. It returns the converted file path and
// a cleanup func that removes the temporary directory. The "Calc MS Excel 2007
// VBA" export filter is named explicitly so macros are preserved (a plain xlsx
// conversion would drop VBA, which is one of this tool's main outputs).
func convertXLSB(src string) (string, func(), error) {
	soffice := findSoffice()
	if soffice == "" {
		return "", nil, fmt.Errorf(`.xlsb は OOXML (ZIP+XML) 形式ではないため直接解体できません。
自動変換に使う LibreOffice (soffice) が見つかりませんでした。次のいずれかで対応してください:
  1. LibreOffice をインストールする (macOS: brew install --cask libreoffice)。
     インストール後はこのツールが自動で .xlsm に変換して解体します。
  2. Excel か LibreOffice で手動で .xlsm へ変換してから渡す
     (VBA マクロを残すため .xlsx ではなく .xlsm を選ぶこと)`)
	}

	tmpDir, err := os.MkdirTemp("", "disassemble_excel_xlsb_*")
	if err != nil {
		return "", nil, err
	}
	cleanup := func() { _ = os.RemoveAll(tmpDir) } // 一時ディレクトリの掃除は best-effort

	fmt.Fprintf(os.Stderr, "  xlsb: LibreOffice で .xlsm に変換中 (%s)\n", filepath.Base(soffice))
	cmd := exec.Command(soffice, "--headless",
		"--convert-to", "xlsm:Calc MS Excel 2007 VBA",
		"--outdir", tmpDir, src)
	out, err := cmd.CombinedOutput()
	if err != nil {
		cleanup()
		return "", nil, fmt.Errorf("soffice での xlsb→xlsm 変換に失敗: %w\n%s", err, out)
	}

	converted := filepath.Join(tmpDir,
		strings.TrimSuffix(filepath.Base(src), filepath.Ext(src))+".xlsm")
	if _, err := os.Stat(converted); err != nil {
		cleanup()
		return "", nil, fmt.Errorf("変換後の .xlsm が見つかりません: %s\nsoffice 出力:\n%s", converted, out)
	}
	return converted, cleanup, nil
}

// findSoffice locates the LibreOffice CLI on PATH or in the macOS default
// install location. Returns "" if not found.
func findSoffice() string {
	for _, name := range []string{"soffice", "libreoffice"} {
		if p, err := exec.LookPath(name); err == nil {
			return p
		}
	}
	for _, p := range []string{
		"/Applications/LibreOffice.app/Contents/MacOS/soffice",
	} {
		if _, err := os.Stat(p); err == nil {
			return p
		}
	}
	return ""
}

func extractVBAFromZip(zr *zip.Reader) ([]Module, error) {
	rc := openZip(zr, "xl/vbaProject.bin")
	if rc == nil {
		return nil, fmt.Errorf("no vbaProject.bin (not a macro workbook)")
	}
	defer func() { _ = rc.Close() }()
	bin, err := io.ReadAll(rc)
	if err != nil {
		return nil, err
	}
	return ExtractVBA(bin)
}

func prepareOutDir(dir string, force bool) error {
	info, err := os.Stat(dir)
	if err == nil {
		if !info.IsDir() {
			return fmt.Errorf("%s exists and is not a directory", dir)
		}
		entries, _ := os.ReadDir(dir)
		if len(entries) > 0 && !force {
			return fmt.Errorf("%s already exists and is not empty (use -f to overwrite)", dir)
		}
	}
	return os.MkdirAll(dir, 0o755)
}

// isOLE2 reports whether the file starts with the OLE2 (CFB) signature: what an encrypted .xlsx/.xlsm or a legacy .xls
// is stored in.
func isOLE2(path string) bool {
	f, err := os.Open(path)
	if err != nil {
		return false
	}
	defer func() { _ = f.Close() }()
	var sig [8]byte
	if _, err := io.ReadFull(f, sig[:]); err != nil {
		return false
	}
	return sig == [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
}

func fileSHA256(path string) string {
	f, err := os.Open(path)
	if err != nil {
		return ""
	}
	defer func() { _ = f.Close() }()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return ""
	}
	return hex.EncodeToString(h.Sum(nil))
}

func lineCount(s string) int {
	if s == "" {
		return 0
	}
	return strings.Count(s, "\n") + 1
}

func slimNote(slim bool) string {
	if slim {
		return "  (+slim)"
	}
	return ""
}
