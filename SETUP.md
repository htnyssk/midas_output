# 引き継ぎ方（別のパソコンで使えるようにする手順）

前提: Windows。Git と Python 3.12 以降が入っていること。PDF 出力が要る場合は Word も。

## A. ソースから動かす場合

### 1. GitHub から取得

```bash
git clone <リポジトリのURL> midas_output
cd midas_output
```

### 2. ライブラリを入れる

```bash
python -m pip install -r requirements.txt
```

`ModuleNotFoundError: No module named 'yaml'` などが出たら、この手順が済んでいない。

### 3. 画像（EMF）をコピーする

`images/` は GitHub に載っていない（合計 1.5GB 超）。元の PC から、USB メモリ・共有フォルダ・ZIP などで
`midas_output\images\` に入れる。

### 4. 設定を確認する

`config.yaml` を開き、案件に合わせて次を確認する。

- `font_family`、余白、画像サイズ
- `sort_order_floors`（階の並び）
- `title_map_chapters`（章タイトル）

各項目の意味はファイル内のコメントを参照。

### 5. 実行する

```bash
python app.py
```

メニューで `3`（リスト作成 → Word 作成）を選ぶ。個別に実行するなら次のとおり。

```bash
python make_list.py
python main.py
```

手順は「`make_list.py` → （`list.xlsx` を Excel で手直し）→ `main.py`」。
出力は `output.docx`（と、Word があれば `output.pdf`）。

## B. 配布用の exe を作る場合

```bash
build.bat
```

`dist\midas_output\` に出来上がる。配布するときは **`midas_output` フォルダごと**（`_internal` も含めて）渡す。
`dist` という名前の親フォルダは不要。

利用者側は、同じフォルダの `images\` に EMF を入れ、`config.yaml` を編集して、`midas_output.exe` を起動する。

## C. 動作確認

1. `images\` に EMF が入っていること。
2. `python app.py` → `1`（リスト作成）。`PNG\` と `list.xlsx` ができ、最後に `Success` が出る。
3. `list.xlsx` を開き、章タイトルが「その他」になっている行が無いこと。ある場合は、`config.yaml` の `filename_rules` にルールを追加する。
4. `python app.py` → `2`（Word 作成）。`output.docx` ができる。

## D. 詰まったとき

| 症状 | 対処 |
|---|---|
| `No module named ...` | `python -m pip install -r requirements.txt` |
| `Configuration file 'config.yaml' not found` | `config.yaml` が実行ファイルと同じフォルダにあるか確認 |
| 「list.xlsx が Excel で開かれています」 | Excel を閉じて再実行 |
| 章タイトルが「その他」 | ファイル名がルールに合っていない。`filename_rules` を確認 |
| PDF が出ない | Word が入っていない。Word 出力（`.docx`）までは動く |
| 変換が遅い・メモリ不足 | 初回のみ遅い。変換済みの `PNG\` は再利用される |

## E. GitHub 側の作業（引き継ぐ側がまだリポジトリを持っていない場合）

元の PC で、次を行う。

1. GitHub で空のリポジトリを作る（**Private** 推奨。README や .gitignore は追加しない）。
2. リポジトリの URL を控える。
3. 元の PC で次を実行する。

```bash
git remote add origin <リポジトリのURL>
git push -u origin main
```
