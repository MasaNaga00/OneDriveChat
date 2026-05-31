# run_pipeline 実行ファイル版 セットアップ＆使用方法

Python をインストールしていない PC でも、配布された **`run_pipeline.exe`** で
案件資料の Markdown 変換とアップロードリスト生成ができる。

このドキュメントは2部構成:
- **A. 利用する人向け**（exe を受け取って使うだけの人）
- **B. ビルドする人向け**（exe を作って配布する人）

---

# A. 利用する人向け

## A-1. 配布物の中身

受け取ったフォルダには次が入っている（フォルダごと任意の場所に置いてよい）。

```
run_pipeline_app/
├─ run_pipeline.exe        ← これを実行する
├─ lib/                    ← 動作に必要な部品（消さない・触らない）
├─ .env                    ← 設定ファイル（自分の環境に合わせて編集）
├─ projects.json           ← 案件の登録（自分の案件に合わせて編集）
└─ （実行すると upload_list.md などがここに出る）
```

`lib` フォルダと exe は**必ず同じフォルダに置いたまま**にする。exe だけ別の場所に
コピーしても動かない。

## A-2. 最初の設定

### .env を編集

メモ帳などで `.env` を開き、自分の環境に合わせる。Dify を使わないなら
Dify 関連は空のままでよい（その場合は後述の `--no-dify` を付けて実行）。

```
# Dify を使う場合のみ
DIFY_API_KEY=app-xxxxxxxxxxxx
DIFY_BASE_URL=https://（社内DifyのURL）/v1

# 検索対象から外すフォルダ・ファイルの接頭辞（カンマ区切り）
EXCLUDE_PREFIXES=_,▪️,■,~,.

# 社内DifyでSSL証明書エラーが出る場合（通常は空でよい）
DIFY_CA_BUNDLE=
DIFY_USE_TRUSTSTORE=true
```

### projects.json を編集

処理したい案件を登録する。フォルダの場所はフルパスで指定。

```
{
  "projects": [
    {
      "name": "案件A",
      "source": "C:/Users/自分/Documents/案件A/資料",
      "out": "C:/Users/自分/CopilotKnowledge/案件A",
      "onedrive_folder": "案件A"
    }
  ]
}
```

- `source`: 元資料のフォルダ
- `out`: 変換した Markdown の出力先
- `onedrive_folder`: アップロード先の OneDrive フォルダ名

## A-3. 実行する

エクスプローラーのアドレスバーに `cmd` と打って Enter（そのフォルダで
コマンドプロンプトが開く）。または PowerShell でフォルダに移動して実行。

```
run_pipeline.exe                  通常実行（全案件を差分処理）
run_pipeline.exe --dry-run        何が処理されるか確認だけ
run_pipeline.exe --no-dify        Difyを使わず変換だけ
run_pipeline.exe --only 案件A      特定案件だけ
run_pipeline.exe --force          全件やり直し
```

**最初は `run_pipeline.exe --dry-run --no-dify` で試す**のがおすすめ。
何も変更せず、対象ファイルや除外が想定通りか確認できる。

## A-4. 実行後

exe と同じフォルダに `upload_list.md` ができる。これを開くと、
案件ごとに「OneDrive に上げるファイル / 消すファイル」がチェックリストで並ぶ。
これを見ながら、ブラウザで OneDrive にアップロード・削除する。

更新がなければ「更新はありません」と出て、作業不要。

## A-5. よくあるトラブル

| 症状 | 対処 |
|---|---|
| 起動した瞬間に黒い窓が一瞬で消える | コマンドプロンプトから起動してエラー文言を確認 |
| 「projects.json が見つかりません」 | exe と同じフォルダに projects.json があるか確認 |
| Dify でSSL/証明書エラー | `.env` の DIFY_USE_TRUSTSTORE=true を確認（社内CA利用時） |
| ウイルス対策ソフトに止められる | IT部門に申請（自作 exe が誤検知されることがある） |
| 文字化けする | コマンドプロンプトを使う。古い環境なら `chcp 65001` を先に実行 |

---

# B. ビルドする人向け

## B-1. 前提

- **配布先と同じ Windows 上でビルドする**（cx_Freeze はクロスビルド不可。
  Windows向け exe は Windows で作る）。
- Python 3.10 以上（truststore が 3.10+ を要求）。
- ビルド用に、仮想環境(venv)を使うことを強く推奨（余計なパッケージを
  exe に巻き込まないため）。

## B-2. 必要ファイル

ビルドフォルダに次を揃える。

```
run_pipeline.py
preprocess_for_copilot.py
enrich_with_dify.py
setup.py
```

（`.env` と `projects.json` はビルドには不要。配布時に手で同梱する。）

## B-3. ライブラリのインストール

```
python -m venv venv
venv\Scripts\activate

pip install python-docx mammoth python-pptx pymupdf pandas openpyxl ^
            extract-msg beautifulsoup4 markdownify requests ^
            python-dotenv truststore cx_Freeze
```

## B-4. ビルド

```
python setup.py build
```

成功すると `build\exe.win-amd64-3.xx\` というフォルダができ、その中に
`run_pipeline.exe` と `lib\` 一式が生成される。

## B-5. 配布物の準備

`build\exe.win-amd64-3.xx\` の中身をそのまま配布フォルダにする。加えて、
利用者が設定できるよう次の2つを同梱する。

- `.env`（`.env.example` をコピーし、社内の既定値を入れておくと親切）
- `projects.json`（`projects.example.json` をコピーしたもの）

最終的な配布フォルダ:

```
run_pipeline_app/
├─ run_pipeline.exe
├─ lib/  （build で生成された全部）
├─ .env
├─ projects.json
└─ README_exe.md  （このA部分だけ抜粋して同梱すると親切）
```

これを zip で配布。利用者は展開して A の手順に従う。

## B-6. ビルドで困ったとき

| 症状 | 対処 |
|---|---|
| exe 実行時に `ModuleNotFoundError` | setup.py の `packages` にそのモジュール名を追加して再ビルド |
| pandas / fitz 関連でエラー | setup.py の `zip_exclude_packages` にそのパッケージを追加 |
| exe が巨大すぎる | venv でビルドしているか確認。`excludes` に不要な大物を追加 |
| 起動が遅い | 初回起動は lib 展開で時間がかかることがある。2回目以降は速い |

## B-7. 仕組みメモ

- `run_pipeline.py` は frozen 実行時、`sys.executable`（=exe）の場所を基準に
  `.env` と `projects.json` を読む。だから利用者は exe と同じフォルダに
  これらを置けばよい。
- 出力（upload_list.md など）も exe と同じフォルダに生成される。
- truststore を同梱しているので、社内CA入りの Windows 証明書ストアを使った
  HTTPS 検証ができる（Dify が社内ホストでも証明書エラーになりにくい）。
