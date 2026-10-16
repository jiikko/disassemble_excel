# disassemble_excel

Excel ワークブック (`.xlsx` / `.xlsm` / `.xlsb`) を diff・grep できるプレーンテキスト群に解体する CLI。使い方・出力レイアウト・仕組み・制限事項の正本は README.md。

目的は「Excel を解体し、プログラムとして diff しやすくする」こと。機能を足すかはこれで決める:
ロジックと作り手の意図 (数式・名前・VBA・入力規則・テーブル・コメント・隠しシート) は取り込み、見た目 (書式・結合・条件付き書式・表示形式) は取り込まない。
壊れた・想定外の XML でも落ちず、推測した値を出さない (読めない部品は空にする)。
本体は excelize を使わない (テストの fixture 生成だけが使う)。

## ファイルの地図

- `main.go` (root) — `cli.Main()` を呼ぶだけ。`go install github.com/jiikko/disassemble_excel@latest` で名前どおりのバイナリにするため root に置く
- `internal/cli/main.go` — フラグ解析と全体の処理フロー (`Main`)。`.xlsb` は `convertXLSB` が LibreOffice で .xlsm に変換してから解体する
- `internal/cli/model.go` — `Cell` などの共有データ型
- `internal/cli/sheets.go` — ワークシート XML を直接パースしてセル (数式・値・型) を読む。共有数式の子セルは親セル (si ごと) から展開する。同じ走査でセル以外の部品も `sheetParts` に拾う。`_xHHHH_` の復号 (`unescapeOOXML`) もここ
- `internal/cli/sharedformula.go` — 共有数式の子セル用に、親の数式の相対参照をずらす (`shiftFormula`)。excelize の `GetCellFormula` はシート名付き参照をずらさないので使わない (#2)
- `internal/cli/objects.go` — シート上のマクロ割り当てオブジェクト (画像・図形・ボタン) の抽出
- `internal/cli/writer.go` — 出力ファイル (`cells.tsv` / `values.csv` / `manifest.json` / 出力先 README 等) の書き出し
- `internal/cli/workbook.go` — `xl/workbook.xml` からシート一覧・表示状態・定義名 (hidden と scope) を読む。rels の解決 (`readRels` / `sheetRels`) もここ
- `internal/cli/sheetextras.go` — セル以外のシートの部品 (入力規則・ハイパーリンク・テーブル・コメント) と、その TSV の書き出し。見た目 (書式・結合・条件付き書式) は扱わない
- `internal/cli/vba.go` — VBA マクロ抽出のパイプライン (mscfb で OLE2 を開き、`ovba` で解凍・解析)
- `internal/ovba/` — [MS-OVBA] の最小実装 (純 Go)。`Decompress` (CompressedContainer 展開) と `ParseDir` (dir ストリーム解析)。仕様の一次情報は `internal/ovba/decompress.go` のパッケージ doc

## 入口

- `go install github.com/jiikko/disassemble_excel@latest` / 手元のソースからは `make build`

## ビルド・テスト

- `make lint` / `make test` (CI の `.github/workflows/ci.yml` も同じ入口)
- E2E: `e2e_test.go` が実バイナリを建てて `testdata/sample.xlsm` を解体し、出力を `testdata/golden/sample/` と全ファイル比較する (元ファイルの絶対パスだけ正規化)。出力は入力だけで決まる: 抽出日時のような実行ごとに変わる値を出さない (`TestE2EDeterministic` が固定する)。
  出力を意図して変えたら `go test -run TestE2E -update .` で golden を作り直し、diff を読んでから commit する
- fixture は `go test -run TestE2E -regen -update .` で作り直す (生成は `fixture_test.go`。`testdata/vbaProject.bin` は excelize のテストデータ。BSD-3-Clause)。
  excelize が書けない形 (x14 の入力規則・隠し定義名・`_x` エスケープ・テーブルの計算列) は保存後の XML を差し替えて作る。
  excelize は数式を計算しないので、fixture の数式セルのキャッシュ値は空

## 詳しくは

- README.md の「機能」「使い方」(フラグ表)「出力レイアウト」「仕組み」「バイナリ形式 `.xlsb` のワークブック」「制限事項」の各節
