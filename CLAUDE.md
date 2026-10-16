# disassemble_excel

Excel ワークブック (`.xlsx` / `.xlsm` / `.xlsb`) を diff・grep できるプレーンテキスト群に解体する CLI。使い方・出力レイアウト・仕組み・制限事項の正本は README.md。

## ファイルの地図

- `main.go` (root) — `cli.Main()` を呼ぶだけ。`go install github.com/jiikko/disassemble_excel@latest` で名前どおりのバイナリにするため root に置く
- `internal/cli/main.go` — フラグ解析と全体の処理フロー (`Main`)。`.xlsb` は `convertXLSB` が LibreOffice で .xlsm に変換してから解体する
- `internal/cli/model.go` — `Cell` などの共有データ型
- `internal/cli/sheets.go` — ワークシート XML を直接パースしてセル (数式・値・型) を読む。共有数式・配列数式は excelize で展開
- `internal/cli/objects.go` — シート上のマクロ割り当てオブジェクト (画像・図形・ボタン) の抽出
- `internal/cli/writer.go` — 出力ファイル (`cells.tsv` / `values.csv` / `manifest.json` / 出力先 README 等) の書き出し
- `internal/cli/defnames.go` — 定義名 (名前付き範囲・LAMBDA 引数) の抽出
- `internal/cli/vba.go` — VBA マクロ抽出のパイプライン (mscfb で OLE2 を開き、`ovba` で解凍・解析)
- `internal/ovba/` — [MS-OVBA] の最小実装 (純 Go)。`Decompress` (CompressedContainer 展開) と `ParseDir` (dir ストリーム解析)。仕様の一次情報は `internal/ovba/decompress.go` のパッケージ doc

## 入口

- `go install github.com/jiikko/disassemble_excel@latest` / 手元のソースからは `make build`

## ビルド・テスト

- `make lint` / `make test` (CI の `.github/workflows/ci.yml` も同じ入口)
- E2E: `e2e_test.go` が実バイナリを建てて `testdata/sample.xlsm` を解体し、出力を `testdata/golden/sample/` と全ファイル比較する (抽出日時と絶対パスだけ正規化)。
  出力を意図して変えたら `go test -run TestE2E -update .` で golden を作り直し、diff を読んでから commit する
- fixture は `go run ./testdata/gen` で作り直す (`vbaProject.bin` は excelize のテストデータ。BSD-3-Clause)。
  excelize は数式を計算しないので、fixture の数式セルのキャッシュ値は空

## 詳しくは

- README.md の「機能」「使い方」(フラグ表)「出力レイアウト」「仕組み」「バイナリ形式 `.xlsb` のワークブック」「制限事項」の各節
