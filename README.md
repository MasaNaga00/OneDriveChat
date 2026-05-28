# 案件ナレッジ前処理パイプライン 使用方法

ローカルの案件資料（Word / PowerPoint / Excel / PDF / Outlookメール / テキスト / HTML）を、
Microsoft 365 Copilot が検索しやすい Markdown に変換し、案件ごとに整理するためのツール群。
更新のあったファイルだけを処理し、OneDrive へのアップロード対象を一覧化する。

---

## 1. これは何をするものか

```
[ローカルの案件フォルダ]
      │  ① 変換（更新分のみ）
      ▼
[Markdown（出典・メタ情報つき）]
      │  ② Dify でメタ情報（分類・キーワード・要約）を自動付与（任意）
      ▼
[アップロード対象リストを出力]
      │  ③ OneDrive へアップロード（当面は手動、将来 Selenium 自動化）
      ▼
[案件ごとの Copilot エージェントが参照]
```

特徴:

- **対応形式**: docx / pptx / xlsx / pdf / msg / txt / html（htm）
- **差分処理**: 前回から変わっていないファイルは変換も Dify 送信もしない（トークン節約）
- **案件ごとに独立管理**: フォルダの場所がバラバラでも個別指定できる
- **出典表記**: 各 Markdown に元ファイル名を埋め込み、Copilot が根拠を示せる
- **メール引用の重複除去**: 引用チェーンを落として新規本文だけ残す
- **除外ルール**: 特定の接頭辞（_ ▪️ など）で始まるフォルダ・ファイルを対象外に
- **同名重複の排除**: 同名の pptx と pdf があれば pptx を優先（pdf はスキップ）

---

## 2. ファイル構成

| ファイル | 役割 |
|---|---|
| `run_pipeline.py` | **メイン。** 差分検知・変換・メタ付与・アップロードリスト出力を統合実行 |
| `preprocess_for_copilot.py` | 各形式を Markdown に変換する処理本体（run_pipeline が内部利用） |
| `enrich_with_dify.py` | Dify でメタ情報を付与する処理（run_pipeline が内部利用 / 単体実行も可） |
| `dify_workflow_meta_generator.yml` | Dify にインポートするワークフロー定義 |
| `copilot_prompts.md` | Copilot の定型質問プロンプト集 |
| `projects.example.json` | 案件設定のテンプレート → `projects.json` にコピーして使う |
| `.env.example` | 接続情報・設定のテンプレート → `.env` にコピーして使う |
| `graph_auth.py` / `graph_upload.py` | Graph API 経由アップロード（※ 環境で Graph 不可のため当面は未使用） |

日常的に実行するのは **`run_pipeline.py` だけ**。他は部品か設定。

---

## 3. 初回セットアップ

### 3-1. Python と必要ライブラリ

Python 3.10 以上を推奨。以下をインストールする。

```
pip install python-docx mammoth python-pptx pymupdf pandas openpyxl extract-msg beautifulsoup4 markdownify requests python-dotenv
```

（Dify を使わず変換だけなら最低限 `python-dotenv` 以外の上記でも動く。Dify 連携には `requests` が必要。）

### 3-2. 設定ファイルを用意

`.env.example` を `.env` にコピーし、Dify と除外設定だけ埋める（Graph 関連は使わないので空のままでよい）。

```
# Dify 接続（メタ情報の自動付与に使う。使わないなら空でも可）
DIFY_API_KEY=app-xxxxxxxxxxxxxxxx
DIFY_BASE_URL=https://api.dify.ai/v1
DIFY_INPUT_VAR=text
DIFY_MAX_CHARS=6000

# 除外する接頭辞（カンマ区切り。これで始まるフォルダ・ファイルを対象外に）
EXCLUDE_PREFIXES=_,▪️,■,~,.

# アップロード対象リストの出力先
UPLOAD_LIST_PATH=./upload_list.json
```

### 3-3. 案件を登録

`projects.example.json` を `projects.json` にコピーし、案件ごとに記入する。
フォルダの場所がバラバラでもフルパスで指定できる。

```
{
  "projects": [
    {
      "name": "案件A_新製品ライン",
      "source": "C:/Users/masashi/Documents/案件A/資料",
      "out": "C:/Users/masashi/CopilotKnowledge/案件A",
      "onedrive_folder": "案件A_新製品ライン"
    },
    {
      "name": "案件B_既存改修",
      "source": "D:/共有/案件B資料",
      "out": "C:/Users/masashi/CopilotKnowledge/案件B",
      "onedrive_folder": "案件B_既存改修"
    }
  ]
}
```

| 項目 | 意味 |
|---|---|
| `name` | 案件名（Markdown のメタ情報にも入る） |
| `source` | 元ファイルがあるフォルダ（読み取り専用で扱う） |
| `out` | 変換した Markdown の出力先。ここに台帳 manifest.json も自動生成 |
| `onedrive_folder` | アップロード先の OneDrive フォルダ名（省略時は name を使用） |

---

## 4. 日常の使い方

### 基本の実行

```
python run_pipeline.py
```

これだけで全案件について「更新分のみ変換 → Dify メタ付与 → アップロードリスト出力」が走る。

### よく使うオプション

| コマンド | 用途 |
|---|---|
| `python run_pipeline.py` | 全案件を差分処理（通常運用） |
| `python run_pipeline.py --dry-run` | 何が処理されるか確認だけ（変換もDifyもせずOK） |
| `python run_pipeline.py --only 案件A` | 案件名に「案件A」を含むものだけ処理 |
| `python run_pipeline.py --no-dify` | Dify に送らず変換のみ（API課金なし） |
| `python run_pipeline.py --force` | 台帳を無視して全件再処理 |

### 実行後に生成されるもの

実行のたびに、更新内容が 3 つの形式で出力される（同じ内容）。

| ファイル | 用途 |
|---|---|
| `upload_list.md` | **手動アップロード用。** 人が読むチェックリスト形式 |
| `upload_list.json` | 将来の Selenium 自動化用 |
| `upload_list.csv` | 表計算で確認したいとき用 |

---

## 5. OneDrive へのアップロード（当面は手動）

`upload_list.md` を開くと、案件ごとに「上げるファイル」「消すファイル」がチェックリストで並ぶ。

```
## 案件: 案件A_新製品ライン
アップロード先フォルダ: `案件A_新製品ライン`

### ⬆ アップロード（このファイルを上記フォルダに上げる）
- [ ] `manual.md`
      場所: `C:/Users/masashi/CopilotKnowledge/案件A/manual.md`

### 🗑 削除（上記フォルダから消す）
- [ ] `old_memo.md`
```

手順:

1. ブラウザで OneDrive を開く
2. リストの「アップロード先フォルダ」を開く
3. 「⬆ アップロード」の各ファイルを、記載の「場所」から探してドラッグ＆ドロップ
4. 「🗑 削除」の各ファイルを、OneDrive 上から削除
5. 終わった項目はチェックを入れて記録

更新がなければ「更新はありません」と表示され、作業不要。

---

## 6. Copilot エージェントの設定と質問

### エージェント作成

案件ごとに、その案件の OneDrive フォルダを参照する Copilot エージェントを作る。
（OneDrive にアップした Markdown が Copilot のインデックス対象になる）

### 定型質問の登録

`copilot_prompts.md` に、すぐ使えるプロンプトが入っている。主力は次の 2 つ。

- **① 進捗状況まとめ** … 案件全体の現状を把握
- **② 次にすべきアクションの提案** … 次の打ち手を考える

補助として、③経緯・決定事項の整理（メール履歴が多い案件向け）、④キャッチアップ、
⑤抜け漏れチェックも収録。詳細と全文は `copilot_prompts.md` を参照。

---

## 7. Dify ワークフローの準備（メタ情報を自動付与する場合）

メタ情報（分類・キーワード・要約）を自動で付けたい場合のみ必要。不要なら `--no-dify` で省ける。

1. Dify で「アプリを作成」→「DSL ファイルをインポート」から
   `dify_workflow_meta_generator.yml` を読み込む
2. LLM ノードのモデルが自分のワークスペースで使えるものか確認（必要なら差し替え）
3. システムプロンプトの分類カテゴリ例を、自分の業務に合わせて調整
4. Dify アプリの「API アクセス」で API キー（app- で始まる）を発行し、`.env` の
   `DIFY_API_KEY` に設定

ワークフローの入出力: 入力変数 `text`（本文）→ 出力 `result`（JSON文字列）。
`.env` の `DIFY_INPUT_VAR` と一致させること（既定は text）。

---

## 8. 仕組みの補足

### 差分検知（トークン節約の核心）

各案件の出力フォルダにある `manifest.json` が台帳。元ファイルの内容ハッシュ（SHA-256）を
記録し、前回と一致すれば変換も Dify 送信もスキップする。
ファイルをコピーしただけ・更新時刻だけ変わった、では再処理されない（中身で判定するため）。

### 除外ルール

`EXCLUDE_PREFIXES` の文字で始まる**フォルダ名・ファイル名**は対象外。
例: `_` を指定 → `_作業中/` 配下も `_メモ.docx` も処理されない。

### 同名 pptx / pdf の扱い

同じフォルダに同名（拡張子違い）があると、優先度の高い形式だけ処理する。
優先順: pptx > docx > xlsx > html > pdf > txt > msg。
送付用 PDF は元の PPTX があれば自動でスキップされる。同名がなければ PDF も処理対象。

### メール（.msg）の引用除去

「-----元のメッセージ-----」「差出人:」などを目印に、過去の引用部分を落として
新規本文だけを残す。引用チェーンの重複でナレッジが膨らむのを防ぐ。
ただし新規本文が極端に短い返信（「承知しました」等）は、文脈のため直近の引用を一部残す。

---

## 9. 困ったとき

| 症状 | 対処 |
|---|---|
| `DIFY_API_KEY が未設定` | `.env` にキーを入れる、または `--no-dify` で変換のみ実行 |
| 案件が処理されない | `projects.json` の `source` パスが正しいか、`--dry-run` で確認 |
| 必要なファイルが除外される | `EXCLUDE_PREFIXES` を見直す（接頭辞が広すぎないか） |
| メールの引用が残る／本文が消えすぎる | 社内特有の引用書式の可能性。区切り行のパターンを調整（要相談） |
| Markdown は出るが Copilot がうまく答えない | OneDrive に最新がアップされているか、フォルダ指定が正しいか確認 |
| 全部作り直したい | `--force` で再処理（または該当案件の out フォルダの manifest.json を削除） |

### 推奨する最初の一歩

```
python run_pipeline.py --dry-run --no-dify
```

これで「どの案件の・どのファイルが対象になるか」「除外や重複排除が意図通りか」を、
何も変更せずに確認できる。問題なければオプションを外して本実行する。
