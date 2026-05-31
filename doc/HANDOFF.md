# 開発引き継ぎ書（Copilot 案件ナレッジ前処理パイプライン）

このドキュメントは、本プロジェクトを別の開発者・別のAIアシスタントに引き継ぐためのもの。
コードの読み方だけでなく「なぜそう作ったか（設計判断と制約）」「未完了事項」を重視して記載する。
**まず最初にこの HANDOFF.md を読み、次に README.md → 各スクリプトの順で読むこと。**

---

## 0. プロジェクトの目的（一言で）

社内の「特定業務に関する相談チャットボット」を **Microsoft 365 Copilot** で作る。
そのために、ローカルにある案件資料（Office文書・PDF・Outlookメール等）を
**Copilot が検索しやすい Markdown に前処理**し、案件ごとに OneDrive へ上げて
案件別の Copilot エージェントに参照させる。前処理の精度向上に Python と Dify を使う。

---

## 1. 確定している前提・制約（重要：これを覆さないこと）

これらはユーザーの環境・要望から確定した制約。引き継ぎ先が「もっと良い方法」を
提案する前に、なぜこの形なのかを理解すること。

1. **Copilot は「標準チャットにファイル参照させる」方式**（Copilot Studio の高度な
   カスタマイズはしない）。よって前処理が精度のほぼ全てを決める。
2. **Microsoft Graph API は使えない**。テナント側で Graph アクセスが制限されており、
   Graph Explorer で `/me/drive` すら見られなかった。→ 自動アップロードに Graph は不可。
3. **OneDrive は Explorer 同期が使えず、常に Web ブラウザでアクセス**している環境。
4. アップロードは当面 **手動**。将来 **Selenium でブラウザ操作**して自動化する予定
   （ユーザーが Selenium 側を別途実装する。本パイプラインは「何を上げるべきか」の
   リストを渡す役割）。
5. **案件ごとにフォルダ・エージェントを分ける**運用。案件フォルダの場所はバラバラ
   （共通ルート無し）。→ projects.json で個別にフルパス指定する方式にした。
6. **メール(.msg)は引用チェーンが多い**（過去のやり取りを引用して保存しているため
   重複が多い）。→ 引用除去処理を入れた。
7. **同じフォルダに同名の pptx と pdf がある**ことが多い（送付用に PDF 化したもの）。
   → pptx を優先して pdf はスキップ。
8. 開発言語は **Python**。ユーザーは Python と Dify が使える環境。Tableau も使う。
9. 最終的に **cx_Freeze で実行ファイル化**し、Python 未導入の同僚にも配る。
10. 利用者は **チームで共有**。メール履歴が多い案件が中心。

---

## 2. 成果物一覧と役割

| ファイル | 役割 | 状態 |
|---|---|---|
| `run_pipeline.py` | **メイン。** 差分検知・変換・Difyメタ付与・アップロードリスト出力を統合。案件ループ・除外・重複排除を担う | 完成・動作確認済 |
| `preprocess_for_copilot.py` | 各形式→Markdown変換の本体。出典表記・メール引用除去・分割もここ | 完成・動作確認済 |
| `enrich_with_dify.py` | Difyでメタ情報付与。TLS証明書の二段構えもここ | 完成・動作確認済 |
| `dify_workflow_meta_generator.yml` | Difyにインポートするワークフロー定義（start→LLM→end） | 完成 |
| `copilot_prompts.md` | Copilot定型質問プロンプト集（①進捗 ②次アクション ③経緯 ④キャッチアップ ⑤抜け漏れ） | 完成 |
| `projects.example.json` | 案件設定テンプレート | 完成 |
| `env_sample.txt` / `.env.sample` | .env のお手本（中身は同一。ドット名が開きにくいため txt 版も用意） | 完成 |
| `setup.py` | cx_Freeze ビルド設定 | 完成・Linuxでビルド検証済 |
| `README.md` | 通常版（スクリプト実行）の使用方法 | 完成 |
| `README_exe.md` | 実行ファイル版の使用方法（利用者向け＋ビルド者向け） | 完成 |
| `graph_auth.py` / `graph_upload.py` | Graph API経由アップロード | **現在未使用**（制約2のため）。将来 Graph が解禁されたら復活可能。残してある |

---

## 3. 全体のデータフロー

```
[ローカルの案件フォルダ群（場所バラバラ）]
      │  run_pipeline.py が projects.json を読んで案件ごとに処理
      │
      │  ① 除外（_ ▪️ 等で始まるフォルダ/ファイル）
      │  ② 同名 pptx/pdf は pptx 優先で重複排除
      │  ③ 差分検知（manifest.json のSHA-256ハッシュ照合。変更分のみ処理）
      │  ④ 形式変換（preprocess_for_copilot）→ Markdown化
      │       - 各 .md に出典表記（先頭・末尾にファイル名）
      │       - .msg は引用除去（新規本文だけ残す）
      │       - メタ枠（title/source/date/category/keywords/project）
      │  ⑤ Difyメタ付与（enrich_with_dify、--no-dify でスキップ可）
      │       - category/keywords/summary を生成しフロントマターに書き戻し
      ▼
[案件ごとの out フォルダに Markdown + manifest.json]
      │  ⑥ upload_list.{md,json,csv} を出力（更新分のみ）
      ▼
[手動 or 将来Selenium で OneDrive へアップ]
      ▼
[案件ごとの Copilot エージェントが参照 → copilot_prompts.md の定型質問で相談]
```

---

## 4. 主要な設計判断（なぜそうしたか）

引き継ぎ先が経緯を知らずに「改善」して壊さないよう、判断の理由を残す。

### 4-1. 元ファイルを直接上げず、Markdownに変換してから上げる
Copilot のインデックスは形式によって抽出精度が落ちる（特に .msg / 複雑なExcel /
レイアウト崩れPDF）。読みやすいMarkdownに正規化することで検索ヒット率と回答精度が上がる。

### 4-2. Graph の解析テキスト上限4MBを意識した分割
Graph(Microsoft Search)のインデックスは1アイテム約4MBで打ち切られる。長文は
見出し単位で分割して後半切り捨てを防ぐ（SPLIT_THRESHOLD_CHARS）。
※ 現状テキストPDF中心・数十ファイル規模なので分割が走ることは稀。

### 4-3. 出典表記は「ファイル名のテキスト埋め込み」方式
ユーザー要望は「どのファイルが根拠か分かれば十分」（クリックリンクは不要）。
各 .md の先頭・末尾にファイル名を埋め込み、プロンプトで「根拠ファイル名を書け」と
指示することで Copilot がそれを拾う。Markdown化で元ファイルとの紐付けが切れる問題への対処。

### 4-4. 差分検知はハッシュ方式（更新時刻ではなく）
Difyのトークン節約が目的。更新時刻はコピー・同期でズレるので、内容のSHA-256で
「中身が同じなら処理しない」を担保。manifest.json に記録。再実行に強い。

### 4-5. Dify連携は「単一 result キーに JSON文字列」構成
ユーザーが普段使っている構成に合わせた。ワークフローは入力 text → 出力 result。
result の中身は {"category","keywords","summary"} の JSON文字列。
enrich側はコードフェンスや前置き付きでもパースできるよう保険を入れてある。

### 4-6. TLS証明書は二段構え（ファイル優先→truststoreでWindowsストア）
社内セルフホストDify / 企業プロキシで証明書検証が失敗する対策。
DIFY_CA_BUNDLE 指定があれば使い、無ければ truststore で OS の証明書ストアに委譲。
truststore.inject_into_ssl() は ssl import より前に呼ぶ必要があるため、
スクリプト冒頭で実行している（順序が重要）。

### 4-7. 実行ファイル化で APP_DIR 基準にした
cx_Freeze の frozen 実行時は sys.executable（exe）の場所、通常時は __file__ の場所を
基準に .env / projects.json を読む。利用者がどこから起動しても exe の隣の設定を読める。

### 4-8. アップロードリストを3形式で出力
upload_list.md（手動作業用・人が読む）/ .json（Selenium用）/ .csv（確認用）。
今は手動運用で .md を使い、Selenium 完成後は .json に移行。同じ差分情報を共有するので
二重管理にならない。削除されたファイルは action="remove" として記録済み。

---

## 5. 設定ファイルの要点

### .env（env_sample.txt 参照）
実際に使う環境変数:
- `DIFY_API_KEY` / `DIFY_BASE_URL` / `DIFY_INPUT_VAR`(既定text) / `DIFY_MAX_CHARS`(既定6000)
- `DIFY_CA_BUNDLE`(空可) / `DIFY_USE_TRUSTSTORE`(既定true)
- `EXCLUDE_PREFIXES`（既定 `_,▪️,■,~,.`）
- `PROJECTS_FILE` / `UPLOAD_LIST_PATH`（未指定なら exe/スクリプトの隣）
- `LOCAL_MARKDOWN_DIR`（enrich単体実行時のみ。通常運用では不要）

### projects.json（projects.example.json 参照）
```
{"projects":[
  {"name":"案件名","source":"元フォルダ(フルパス)","out":"出力先(フルパス)","onedrive_folder":"OneDrive上の上げ先フォルダ名"}
]}
```
案件ごとに out フォルダ内に manifest.json が自動生成される（案件ごとに独立）。

---

## 6. 動作確認済みの範囲

この環境（Linux）で以下を実機検証済み:
- 全形式（docx/pptx/xlsx/pdf/txt/html/msgロジック）の変換
- 差分検知（新規・変更・変更なし・削除）
- 除外（フォルダ・ファイル両方）
- 同名 pptx/pdf の重複排除
- メール引用除去（3パターン＋短文時の文脈保持）
- Dify連携のフロントマター往復・JSON抽出（実APIは未接続。ロジックのみ検証）
- TLS証明書の二段構え（4パターン＋実HTTPS検証）
- cx_Freeze ビルド → exe起動 → 変換 → upload_list生成（Linux版で確認）

---

## 7. 未完了・未検証・今後の課題（引き継ぎ先が取り組む領域）

### 7-1. 【最優先】Selenium アップロード自動化（未着手）
- 本パイプラインが出力する `upload_list.json` を読んで、OneDrive(Web)に
  ブラウザ操作でアップロード/削除する Python を作る。
- json の各 task: `{action: upload|remove, local_path, onedrive_folder, file_name}`
- ログイン（社内SSO等）の扱い、案件名→OneDrive実フォルダの対応づけが論点。
- ユーザーが自分で実装予定だが、引き継ぎ先が支援する可能性あり。

### 7-2. Dify 実APIでの疎通確認（未実施）
- enrich_with_dify は実 Dify に接続したテストをしていない（ロジックのみ検証）。
- ワークフロー(dify_workflow_meta_generator.yml)をインポートし、実際に
  category/keywords/summary が期待通り返るか、プロンプト調整が必要。

### 7-3. cx_Freeze の Windows 実機ビルド（未実施）
- Linux でのビルドは成功したが、**Windows 向け exe は Windows でビルドが必要**
  （クロスビルド不可）。Windows 固有の隠れ依存が出る可能性。
- 出たら setup.py の packages/includes に追加（README_exe.md の B-6 参照）。

### 7-4. メール引用除去パターンの実データ調整（要確認）
- 日本語Outlook前提でパターンを組んだが、社内固有の書式は取りこぼす可能性。
- 実 .msg で「引用が残る/本文が消えすぎる」が出たら preprocess_for_copilot.py の
  `_QUOTE_PATTERNS` を調整。

### 7-5. 案件数スケール時の方針（将来判断）
- 現状は「案件ごとにフォルダ・エージェント」。案件が数十を大きく超えると
  エージェント管理が煩雑に。その場合は project フロントマターで絞る統合方式に
  切り替える選択肢あり（各 .md に project: 案件名 を既に埋めてあるので移行可能）。

### 7-6. Copilot の同時参照ファイル数の制約（運用上の注意）
- 標準チャットでのファイル参照には同時参照数の上限がある。案件あたり数十ファイルを
  超えると全資料を完全網羅できない可能性。プロンプトで重要ファイルを誘導する等で対処。

---

## 8. 引き継ぎ先（GPTベースAI）への申し送り

- **コードは完成度が高く動作確認済み**。むやみにリファクタせず、上記7の未完了領域に集中するとよい。
- **制約（第1章）を最優先**。特に「Graph不可」「アップロードはSelenium/手動」を忘れない。
  Graph を使う提案は環境的に不可。
- 各スクリプトは**部品が関数化**されており、run_pipeline が preprocess と enrich を
  import して再利用する構造。重複実装しないこと。
- ユーザーは Python 中級者。Dify のワークフロー構築経験あり。Tableau利用者。
- 日本語でのやり取りが基本。
