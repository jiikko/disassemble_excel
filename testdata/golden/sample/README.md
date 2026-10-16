# disassemble_excel 出力ディレクトリ

このディレクトリは Excel ワークブックを `disassemble_excel` で分解（解体）した結果です。
中身はすべて行指向のプレーンテキストで、`diff` / `grep` / バージョン管理でそのまま比較できます。
Excel を開かずに、シートのセル数式・キャッシュ値・定義名・VBA マクロのソースを読むためのものです。

## 生成元

| 項目 | 値 |
|------|----|
| 生成ツール | disassemble_excel v0.1.0 (https://github.com/jiikko/disassemble_excel) |
| 元 Excel ファイル（絶対パス） | `<ABS_PATH>` |
| 元 Excel ファイル名 | `sample.xlsm` |
| 元ファイルの SHA-256 | `6a263a539897aff3f857689c65adf937ae8ef54364414f992f054f2244cdc0cb` |
| 抽出日時 | <EXTRACTED_AT> |

> 数式は **抽出のみで、評価（再計算）はしていません**。各セルの `value` 列は Excel が
> 最後にファイルへ保存したキャッシュ値であり、このツールが計算した値ではありません。

## ファイル構成と各ファイルの意味

| パス | 内容 |
|------|------|
| `manifest.json` | 機械可読の目録。元ファイル名・SHA-256・抽出日時・ツールバージョン、シートごとの状態 (visible / hidden / veryHidden)・寸法・セル数・数式数、VBA モジュール一覧、下の TSV の件数 (counts) を持つ |
| `README.md` | このファイル（出力の説明・元ソース・各ファイルの意味） |
| `defined_names.tsv` | 定義名（名前付き範囲・LAMBDA 引数）。列は `name  scope  refers_to  hidden`（hidden は隠し名なら 1） |
| `sheets/<シート名>.cells.tsv` | シートの全セルを 1 セル 1 行で。列は `cell  type  formula  value`、(行,列) でソート・空セルは省略。値中のタブ/改行は `\t` `\n` にエスケープ |
| `sheets/<シート名>.values.csv` | キャッシュ値だけをグリッド（行×列）に並べた目視確認用ビュー |
| `vba/<モジュール名>.bas` / `.cls` | VBA モジュールのソースコード（標準モジュール=`.bas`、クラス/ドキュメントモジュール=`.cls`）。コードを持つモジュールごとに 1 ファイル |
| `vba_index.tsv` | VBA の Sub/Function/Property の索引。列は `module  proc  kind  module_line` |
| `data_validations.tsv` | データの入力規則（入力できる値の制約・リストの出どころ）。列は `sheet  sqref  type  operator  formula1  formula2  allow_blank`。数式は先頭の = なし |
| `hyperlinks.tsv` | ハイパーリンク。列は `sheet  cell  target  location  display`（target は外部、location はブック内の飛び先） |
| `tables.tsv` | テーブル（構造化参照 `Table1[列]` の名前）の範囲と列。1 列 1 行。列は `table  sheet  ref  column  calculated_formula` |
| `comments.tsv` | セルのコメント（メモ）。列は `sheet  cell  author  text` |
| `objects.tsv` | マクロが割り当てられたシート上のオブジェクト（画像・図形・ボタン）。クリック時に実行されるマクロが分かる。列は `sheet  object  kind  anchor  macro`。画像自体は復元しない |

## シート一覧（3 枚）

| シート | 状態 | 範囲 | セル数 | 数式数 | ファイル |
|--------|------|------|-------:|-------:|----------|
| Summary | visible | A1 | 20 | 3 | `sheets/Summary.cells.tsv` |
| Data | visible | A1 | 20 | 8 | `sheets/Data.cells.tsv` |
| Lists | veryHidden | A1 | 3 | 0 | `sheets/Lists.cells.tsv` |

## VBA モジュール一覧（4 個）

| モジュール | 種別 | プロシージャ数 | 行数 | ファイル |
|------------|------|---------------:|-----:|----------|
| ThisWorkbook | class | 0 | 0 | （空モジュール: ファイル化なし） |
| Sheet1 | class | 2 | 8 | `vba/Sheet1.cls` |
| ThisWorkbook1 | class | 0 | 0 | （空モジュール: ファイル化なし） |
| Module1 | standard | 1 | 3 | `vba/Module1.bas` |

## マクロ割り当てオブジェクト一覧（2 個）

シート上の画像・図形・ボタンに登録され、クリック時に実行されるマクロです。

| シート | オブジェクト | 種別 | 位置 | マクロ |
|--------|--------------|------|------|--------|
| Summary | Shape 2 | sp | D2 | `Button1_Click` |
| Summary | _x0000_s1025 | button | D6 | `Button2_Click` |

