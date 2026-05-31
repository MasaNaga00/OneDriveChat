# 全コードバンドル（引き継ぎ用）

本ファイルは全成果物を1つに連結したもの。各ファイルの中身を通読できる。
編集時は個別ファイルを正とすること（これは閲覧用スナップショット）。

収録順:
  1. run_pipeline.py — メイン統合スクリプト
  2. preprocess_for_copilot.py — 形式変換の本体
  3. enrich_with_dify.py — Difyメタ付与＋TLS設定
  4. setup.py — cx_Freezeビルド設定
  5. dify_workflow_meta_generator.yml — Difyワークフロー定義
  6. projects.example.json — 案件設定テンプレート
  7. env_sample.txt — .envお手本
  8. graph_auth.py — 【未使用】Graph認証
  9. graph_upload.py — 【未使用】Graphアップロード

======================================================================


======================================================================
## ファイル: run_pipeline.py
## 役割: メイン統合スクリプト
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_pipeline.py — 差分検知つきパイプライン統合実行
==================================================
ローカルフォルダを監視し、「新規・変更されたファイルだけ」を
  1. Markdown 変換 (preprocess_for_copilot の変換関数を再利用)
  2. Dify でメタ情報付与 (enrich_with_dify の関数を再利用)
の順に処理する。変更のないファイルは一切 Dify に送らない(トークン節約)。

処理台帳 manifest.json に各元ファイルの SHA-256 を記録し、
前回と一致すればスキップ。削除された元ファイルに対応する
Markdown も掃除する。

設定は projects.json(案件ごとの場所) と .env(Dify共通設定) から読む:
    projects.json: 各案件の source(元フォルダ) / out(出力先) / name(案件名)
                   フォルダの場所がバラバラでもフルパスで個別指定できる。
    .env:          DIFY_API_KEY / DIFY_BASE_URL / DIFY_INPUT_VAR / DIFY_MAX_CHARS

各案件は独立した台帳(manifest)を持つので差分管理が混ざらない。
出力Markdownのフロントマターには project: 案件名 が自動で入る。

使い方:
    python run_pipeline.py                  # projects.json の全案件を差分処理
    python run_pipeline.py --only 案件A      # 案件名に"案件A"を含むものだけ
    python run_pipeline.py --force          # 全ファイル再処理(台帳無視)
    python run_pipeline.py --no-dify        # 変換だけ(Difyに送らない)
    python run_pipeline.py --dry-run        # 何を処理するか表示するだけ
    python run_pipeline.py --projects 別設定.json
"""

import os
import sys
import json
import hashlib
import argparse
from pathlib import Path

# 実行ファイル(cx_Freeze)化されている場合は exe の場所、
# 通常の .py 実行なら このファイルの場所を「基準ディレクトリ」とする。
# これにより、どこから起動しても .env / projects.json を exe の隣から読める。
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

# .env 読み込み(基準ディレクトリの .env を明示的に指す)
try:
    from dotenv import load_dotenv
    load_dotenv(APP_DIR / ".env")
except ImportError:
    pass

# 既存スクリプトの部品を再利用(車輪の再発明をしない)
import preprocess_for_copilot as pp
from enrich_with_dify import (
    split_frontmatter, build_frontmatter, call_dify, merge_meta,
)

MANIFEST_VERSION = 1

# このいずれかの文字で始まるフォルダ名は、配下ごとRAG対象外にする。
# .env の EXCLUDE_PREFIXES で上書き可能(カンマ区切り)。
DEFAULT_EXCLUDE_PREFIXES = ["_", "▪️", "■", "・", "~", "."]


def get_exclude_prefixes() -> list:
    raw = os.getenv("EXCLUDE_PREFIXES", "").strip()
    if raw:
        return [p for p in (s.strip() for s in raw.split(",")) if p]
    return DEFAULT_EXCLUDE_PREFIXES


def is_excluded(rel_path: Path, prefixes: list) -> bool:
    """相対パスの途中のフォルダ名、またはファイル名自体が、
    除外プレフィックスで始まれば True。
    例: prefixes=['_'] のとき
        '_作業中/資料.docx'  → フォルダ _作業中 で除外
        'A/_メモ.txt'        → ファイル名 _メモ.txt で除外
        'A/B/正式.docx'      → 除外されない"""
    # フォルダ部分とファイル名の両方を判定対象にする
    for part in rel_path.parts:
        for pre in prefixes:
            if part.startswith(pre):
                return True
    return False


# 同じ basename で両方あった場合に優先する拡張子(左ほど優先)。
# PPTX を PDF より優先: PDFは送付用に固めたもので、PPTXの方が構造を取りやすい。
FORMAT_PRIORITY = [".pptx", ".docx", ".xlsx", ".html", ".htm", ".pdf", ".txt", ".msg"]


def dedup_same_name(files: list) -> tuple:
    """同じフォルダ・同じファイル名(拡張子違い)が複数ある場合、
    FORMAT_PRIORITY に従って1つだけ残す。
    返り値: (残したファイル一覧, 除外したファイル一覧)"""
    # (フォルダ, ベース名) ごとに候補をまとめる
    groups = {}
    for p in files:
        gkey = (str(p.parent), p.stem)
        groups.setdefault(gkey, []).append(p)

    kept, dropped = [], []
    for gkey, cands in groups.items():
        if len(cands) == 1:
            kept.append(cands[0])
            continue
        # 優先順位でソート(リストにない拡張子は最後尾)
        def rank(path):
            ext = path.suffix.lower()
            return FORMAT_PRIORITY.index(ext) if ext in FORMAT_PRIORITY else 999
        cands_sorted = sorted(cands, key=rank)
        kept.append(cands_sorted[0])
        dropped.extend(cands_sorted[1:])
    return kept, dropped


# ---------------------------------------------------------------------------
# ハッシュ & 台帳
# ---------------------------------------------------------------------------
def file_hash(path: Path) -> str:
    """ファイル内容の SHA-256。大きいファイルも省メモリで計算。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def load_manifest(path: Path) -> dict:
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return data.get("files", {})
        except (json.JSONDecodeError, KeyError):
            print("  ⚠ manifest が壊れています。全件再処理します。")
    return {}


def save_manifest(path: Path, files: dict):
    path.write_text(json.dumps(
        {"version": MANIFEST_VERSION, "files": files},
        ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# 1ファイルの変換 (preprocess の関数を組み合わせて Markdown を書き出す)
# ---------------------------------------------------------------------------
def convert_one(src: Path, out_dir: Path, project: str = "") -> list[Path]:
    """1つの元ファイルを Markdown 化して書き出す。出力パス一覧を返す。
    preprocess_for_copilot の変換ロジック(分割・出典表記)を踏襲。
    project を渡すとフロントマターに project 行を追加する。"""
    body = pp.CONVERTERS[src.suffix.lower()](src)
    fm = pp.build_frontmatter(src)
    if project:
        # フロントマター末尾の閉じ "---" の直前に project 行を挿入する
        idx = fm.rstrip().rfind("---")
        fm = fm[:idx] + f"project: {project}\n" + fm[idx:]
    base = pp.safe_name(src.stem)
    outputs = []

    if len(fm + body) > pp.SPLIT_THRESHOLD_CHARS:
        chunks = pp.split_by_heading(body, pp.SPLIT_THRESHOLD_CHARS)
        total = len(chunks)
        for idx, chunk in enumerate(chunks, 1):
            out = out_dir / f"{base}_part{idx}.md"
            cbody = pp.source_header(src, idx, total) + chunk + pp.source_footer(src)
            out.write_text(fm + cbody, encoding="utf-8")
            outputs.append(out)
    else:
        out = out_dir / f"{base}.md"
        full = pp.source_header(src) + body + pp.source_footer(src)
        out.write_text(fm + full, encoding="utf-8")
        outputs.append(out)
    return outputs


# ---------------------------------------------------------------------------
# 1ファイルの Dify メタ付与 (enrich の関数を再利用)
# ---------------------------------------------------------------------------
def enrich_one(md_path: Path, api_key: str, base_url: str,
               input_var: str, max_chars: int) -> bool:
    """Markdown にメタ情報を付与。付与したら True。"""
    content = md_path.read_text(encoding="utf-8")
    fm, body = split_frontmatter(content)
    # 既に付いていればスキップ(再処理時の二重課金防止の保険)
    if fm.get("category") and fm.get("keywords"):
        return False
    meta = call_dify(body[:max_chars], api_key, base_url, input_var)
    fm = merge_meta(fm, meta)
    md_path.write_text(build_frontmatter(fm) + body, encoding="utf-8")
    return True


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def process_project(proj: dict, opts: dict) -> dict:
    """1案件分を処理して集計を返す。
    proj: {"name","source","out","manifest"(任意)}
    opts: {"force","no_dify","dry_run","api_key","base_url","input_var","max_chars"}
    """
    name = proj.get("name", "(無名)")
    source_dir = Path(proj["source"])
    out_dir = Path(proj["out"])
    manifest_path = Path(proj.get("manifest") or (out_dir / "manifest.json"))

    result = {"name": name, "changed": 0, "unchanged": 0,
              "deleted": 0, "enriched": 0, "error": None,
              "upload": [], "remove": [],
              "onedrive_folder": proj.get("onedrive_folder", name)}

    print(f"\n==== 案件: {name} ====")
    if not source_dir.exists():
        print(f"  ⚠ 元フォルダがありません: {source_dir}  → スキップ")
        result["error"] = "source not found"
        return result
    out_dir.mkdir(parents=True, exist_ok=True)

    use_dify = not opts["no_dify"]
    prefixes = opts["exclude_prefixes"]
    all_files = [p for p in sorted(source_dir.rglob("*"))
                 if p.is_file() and p.suffix.lower() in pp.CONVERTERS]
    # 除外プレフィックスで始まるフォルダ/ファイルを除く
    kept_files, excluded_n = [], 0
    for p in all_files:
        if is_excluded(p.relative_to(source_dir), prefixes):
            excluded_n += 1
        else:
            kept_files.append(p)
    # 同名(拡張子違い)の重複を排除。PPTX優先でPDFを落とす等。
    src_files, dropped = dedup_same_name(kept_files)
    for d in dropped:
        print(f"    (重複スキップ) {d.relative_to(source_dir)} "
              f"… 同名の優先形式を採用")

    manifest = {} if opts["force"] else load_manifest(manifest_path)
    new_manifest = {}
    seen_keys = set()

    print(f"  元: {source_dir}")
    print(f"  出力: {out_dir}")
    print(f"  対象 {len(src_files)} 件 (台帳 {len(manifest)} 件"
          + (f" / 除外 {excluded_n} 件" if excluded_n else "")
          + (f" / 重複 {len(dropped)} 件" if dropped else "") + ")")

    for src in src_files:
        key = str(src.relative_to(source_dir))
        seen_keys.add(key)
        h = file_hash(src)
        prev = manifest.get(key)

        if prev and prev.get("hash") == h and not opts["force"]:
            new_manifest[key] = prev
            result["unchanged"] += 1
            continue

        result["changed"] += 1
        if opts["dry_run"]:
            label = "新規" if not prev else "変更"
            print(f"    [{label}] {key}")
            new_manifest[key] = {"hash": h,
                                 "outputs": prev.get("outputs", []) if prev else []}
            continue

        if prev:
            for old in prev.get("outputs", []):
                Path(old).unlink(missing_ok=True)
        try:
            outputs = convert_one(src, out_dir, project=name)
            out_strs = [str(o) for o in outputs]
            if use_dify:
                for md in outputs:
                    if enrich_one(md, opts["api_key"], opts["base_url"],
                                  opts["input_var"], opts["max_chars"]):
                        result["enriched"] += 1
            new_manifest[key] = {"hash": h, "outputs": out_strs}
            tag = "新規" if not prev else "変更"
            print(f"    ✓ [{tag}] {key} → {len(outputs)}ファイル")
            # Selenium に渡すアップロード対象として記録
            for o in out_strs:
                result["upload"].append(o)
        except Exception as e:
            print(f"    ✗ {key} ({e})")

    # 削除検知
    deleted = [k for k in manifest if k not in seen_keys]
    for key in deleted:
        for old in manifest[key].get("outputs", []):
            if opts["dry_run"]:
                print(f"    [削除] {old}")
            else:
                Path(old).unlink(missing_ok=True)
            # Selenium 側で OneDrive 上からも消せるよう、ファイル名を記録
            result["remove"].append(Path(old).name)
        if not opts["dry_run"]:
            print(f"    🗑 削除: {key}")
    result["deleted"] = len(deleted)

    if not opts["dry_run"]:
        save_manifest(manifest_path, new_manifest)
    return result


def load_projects(path: Path) -> list:
    data = json.loads(path.read_text(encoding="utf-8"))
    projects = data.get("projects", [])
    valid = []
    for p in projects:
        if not p.get("source") or not p.get("out"):
            print(f"  ⚠ source/out が無い案件をスキップ: {p.get('name','(無名)')}")
            continue
        valid.append(p)
    return valid


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--projects", default=os.getenv("PROJECTS_FILE", str(APP_DIR / "projects.json")),
                    help="案件設定JSON(既定 <exeの場所>/projects.json または .env の PROJECTS_FILE)")
    ap.add_argument("--only", help="この案件名だけ処理(部分一致)")
    ap.add_argument("--force", action="store_true", help="台帳を無視して全件再処理")
    ap.add_argument("--no-dify", action="store_true", help="変換のみ。Difyに送らない")
    ap.add_argument("--dry-run", action="store_true", help="判定だけ表示")
    args = ap.parse_args()

    projects_file = Path(args.projects)
    if not projects_file.exists():
        print(f"案件設定が見つかりません: {projects_file}")
        print("projects.example.json を projects.json にコピーして編集してください。")
        sys.exit(1)

    projects = load_projects(projects_file)
    if args.only:
        projects = [p for p in projects if args.only in p.get("name", "")]
    if not projects:
        print("処理対象の案件がありません。")
        sys.exit(0)

    opts = {
        "force": args.force,
        "no_dify": args.no_dify,
        "dry_run": args.dry_run,
        "api_key": os.getenv("DIFY_API_KEY", ""),
        "base_url": os.getenv("DIFY_BASE_URL", "https://api.dify.ai/v1"),
        "input_var": os.getenv("DIFY_INPUT_VAR", "text"),
        "max_chars": int(os.getenv("DIFY_MAX_CHARS", "6000")),
        "exclude_prefixes": get_exclude_prefixes(),
    }
    if not opts["no_dify"] and not opts["api_key"] and not args.dry_run:
        print("DIFY_API_KEY が未設定です。--no-dify で変換のみ実行できます。")
        sys.exit(1)

    print(f"処理する案件: {len(projects)} 件")
    results = [process_project(p, opts) for p in projects]

    # 全案件の総括
    print("\n" + "=" * 40)
    print("全案件サマリー")
    print("=" * 40)
    tot = {"changed": 0, "unchanged": 0, "deleted": 0, "enriched": 0}
    for r in results:
        mark = "⚠" if r["error"] else "・"
        print(f" {mark} {r['name']}: 変更{r['changed']} / "
              f"据置{r['unchanged']} / 削除{r['deleted']}"
              + (f" / Dify{r['enriched']}" if not opts["no_dify"] else ""))
        for k in tot:
            tot[k] += r[k]
    print("-" * 40)
    print(f" 合計: 変更{tot['changed']} / 据置{tot['unchanged']} "
          f"(Dify送信なし=節約) / 削除{tot['deleted']}"
          + (f" / Difyメタ付与{tot['enriched']}ファイル" if not opts["no_dify"] else ""))
    if args.dry_run:
        print("\n(dry-run: 実際の処理はしていません)")
        return

    # ---- Selenium に渡すアップロード指示リストを書き出す ----
    # 更新のあったファイルだけを「ローカルパス → OneDrive上げ先フォルダ」で列挙。
    write_upload_list(results)


def write_upload_list(results: list):
    """更新ファイルのアップロード指示を JSON と CSV で出力する。
    Selenium 側はこれを読んで、記載されたファイルだけを上げればよい。"""
    import datetime
    out_path_json = Path(os.getenv("UPLOAD_LIST_PATH", str(APP_DIR / "upload_list.json")))

    tasks = []
    for r in results:
        if r["error"]:
            continue
        folder = r["onedrive_folder"]
        for local in r["upload"]:
            tasks.append({
                "action": "upload",
                "local_path": str(Path(local).resolve()),
                "onedrive_folder": folder,      # この案件の上げ先(OneDrive側フォルダ名)
                "file_name": Path(local).name,
            })
        for fname in r["remove"]:
            tasks.append({
                "action": "remove",
                "onedrive_folder": folder,
                "file_name": fname,
            })

    payload = {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "total": len(tasks),
        "tasks": tasks,
    }
    out_path_json.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # 人が見て確認しやすいよう CSV も併せて出力
    out_path_csv = out_path_json.with_suffix(".csv")
    lines = ["action,onedrive_folder,file_name,local_path"]
    for t in tasks:
        lines.append(",".join([
            t["action"], t["onedrive_folder"], t["file_name"],
            t.get("local_path", ""),
        ]))
    out_path_csv.write_text("\n".join(lines), encoding="utf-8")

    # ---- 人が手作業するための分かりやすい指示書(Markdown) ----
    md_path = out_path_json.with_suffix(".md")
    md = []
    md.append("# OneDrive 手動更新 作業リスト")
    md.append("")
    md.append(f"生成: {payload['generated_at']}")
    md.append("")
    upn = sum(1 for t in tasks if t["action"] == "upload")
    rmn = sum(1 for t in tasks if t["action"] == "remove")
    if upn == 0 and rmn == 0:
        md.append("**更新はありません。** OneDrive 側の作業は不要です。")
    else:
        md.append(f"このリストの通りに OneDrive を更新してください"
                  f"（アップロード {upn} 件 / 削除 {rmn} 件）。")
        md.append("チェックボックスは作業済みの記録用です。")
        md.append("")
        # 案件ごとにまとめる
        for r in results:
            if r["error"] or (not r["upload"] and not r["remove"]):
                continue
            md.append(f"## 案件: {r['name']}")
            md.append(f"アップロード先フォルダ: `{r['onedrive_folder']}`")
            md.append("")
            if r["upload"]:
                md.append("### ⬆ アップロード（このファイルを上記フォルダに上げる）")
                for local in r["upload"]:
                    p = Path(local)
                    md.append(f"- [ ] `{p.name}`")
                    md.append(f"      場所: `{p.resolve()}`")
                md.append("")
            if r["remove"]:
                md.append("### 🗑 削除（上記フォルダから消す）")
                for fname in r["remove"]:
                    md.append(f"- [ ] `{fname}`")
                md.append("")
    md_path.write_text("\n".join(md), encoding="utf-8")

    print(f"\n📤 更新リストを書き出しました:")
    print(f"   {md_path}    ← 手動作業用（人が読む）")
    print(f"   {out_path_json}  ← Selenium用（アップロード {upn} / 削除 {rmn}）")
    print(f"   {out_path_csv}   ← 確認用CSV")
    if upn == 0 and rmn == 0:
        print("   ※ 更新ファイルなし。アップロード作業は不要です。")


if __name__ == "__main__":
    main()
```

======================================================================
## ファイル: preprocess_for_copilot.py
## 役割: 形式変換の本体
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Microsoft 365 Copilot 向け ナレッジ前処理スクリプト
====================================================
ローカルフォルダ内の docx / pptx / xlsx / pdf / msg / txt / html を、
検索されやすい Markdown に正規化して出力する。

使い方:
    python preprocess_for_copilot.py  入力フォルダ  出力フォルダ

出力された Markdown フォルダを OneDrive(将来 SharePoint) にアップし、
Copilot に参照させる。

前提: 数十ファイル / テキストPDF中心 / スキャンPDFなし
"""

import sys
import re
import datetime
from pathlib import Path

import mammoth                 # docx -> markdown
import fitz                    # PyMuPDF: pdf
import pandas as pd            # xlsx
from pptx import Presentation  # pptx
import extract_msg            # .msg
from bs4 import BeautifulSoup            # html パース
from markdownify import markdownify as _md  # html -> markdown

# Copilot/Graph のインデックス上限(解析テキスト4MB)に対する安全側のしきい値。
# 1ファイルがこれを超えそうなら章/トピック単位に分割する。
SPLIT_THRESHOLD_CHARS = 1_200_000  # 約1.2M文字 (日本語で約3.5MB相当の安全圏)


# ---------------------------------------------------------------------------
# 形式ごとの抽出関数。いずれも本文(Markdown文字列)を返す。
# ---------------------------------------------------------------------------
def conv_docx(path: Path) -> str:
    with open(path, "rb") as f:
        result = mammoth.convert_to_markdown(f)
    return result.value


def conv_pptx(path: Path) -> str:
    prs = Presentation(path)
    blocks = []
    for i, slide in enumerate(prs.slides, 1):
        blocks.append(f"## スライド {i}")
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                blocks.append(shape.text_frame.text.strip())
        # 表があれば拾う
        for shape in slide.shapes:
            if shape.has_table:
                tbl = shape.table
                rows = []
                for r in tbl.rows:
                    rows.append(" | ".join(c.text.strip() for c in r.cells))
                if rows:
                    blocks.append("\n".join(rows))
        # 発表者ノート(意外と重要な情報が入っている)
        if slide.has_notes_slide:
            note = slide.notes_slide.notes_text_frame.text.strip()
            if note:
                blocks.append(f"> ノート: {note}")
    return "\n\n".join(blocks)


def conv_xlsx(path: Path) -> str:
    xls = pd.ExcelFile(path)
    blocks = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if df.empty:
            continue
        blocks.append(f"## シート: {sheet}")
        # 表として読める形(Markdownテーブル)
        blocks.append(df.to_markdown(index=False))
        # 補助として「1行=1文」の文章化も付ける。検索ヒット率が上がる。
        sentences = []
        for _, row in df.iterrows():
            parts = [f"{col}は{val}" for col, val in row.items()
                     if pd.notna(val) and str(val).strip()]
            if parts:
                sentences.append("・" + "、".join(parts) + "。")
        if sentences:
            blocks.append("### 内容(文章化)\n" + "\n".join(sentences))
    return "\n\n".join(blocks)


def conv_pdf(path: Path) -> str:
    doc = fitz.open(path)
    blocks = []
    for i, page in enumerate(doc, 1):
        text = page.get_text("text").strip()
        if text:
            blocks.append(f"## ページ {i}\n\n{text}")
    doc.close()
    body = "\n\n".join(blocks)
    # テキストがほぼ無い = スキャンPDFの可能性。警告だけ出す(今回は対象外想定)。
    if len(body.strip()) < 30:
        body += "\n\n> ⚠ テキストがほとんど抽出できませんでした。スキャンPDFの可能性があります(要OCR)。"
    return body


# 日本語Outlookの引用区切りパターン(これ以降を引用とみなす)
_QUOTE_PATTERNS = [
    r"^-+\s*元のメッセージ\s*-+",        # -----元のメッセージ-----
    r"^-+\s*Original Message\s*-+",        # 英語混在の保険
    r"^-+\s*転送されたメッセージ\s*-+",    # 転送
    r"^_{5,}",                              # ____ の区切り線
    r"^差出人:\s*.+",                       # 差出人: ヘッダ
    r"^From:\s*.+",                         # From: ヘッダ
    r"^送信者:\s*.+",                       # 送信者: の表記ゆれ
]
_QUOTE_COMPILED = [re.compile(p) for p in _QUOTE_PATTERNS]
# 「2024年1月1日(月) 山田 <..@..>:」などの返信定型句
_REPLY_INTRO = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日.*(?:書きました|<.+@.+>).*[:：]?\s*$")


def _find_quote_start(lines):
    for i, line in enumerate(lines):
        s = line.strip()
        if any(c.match(s) for c in _QUOTE_COMPILED) or _REPLY_INTRO.match(s):
            return i
    return None


def strip_quoted(body: str, min_new_chars: int = 20, context_lines: int = 8) -> str:
    """メール本文から引用(過去のやり取り)を除去し新規本文だけ残す。
    引用チェーンの重複を減らすのが目的。
    ただし新規本文が極端に短い場合(「承知しました」等)は、
    直近の引用を少しだけ残して文脈を保つ。"""
    lines = body.splitlines()
    cut = _find_quote_start(lines)
    if cut is None:
        return body.rstrip()

    new_part = "\n".join(lines[:cut]).rstrip()

    if len(new_part.replace("\n", "").strip()) < min_new_chars:
        quoted = lines[cut:]
        context = []
        for j, line in enumerate(quoted):
            if j > 0 and (any(c.match(line.strip()) for c in _QUOTE_COMPILED)
                          or _REPLY_INTRO.match(line.strip())):
                break
            context.append(line)
            if len(context) >= context_lines:
                break
        ctx = "\n".join(context).rstrip()
        return new_part + "\n\n[文脈のため直近の引用を一部保持]\n" + ctx
    return new_part


def conv_msg(path: Path) -> str:
    msg = extract_msg.Message(str(path))
    header = []
    if msg.subject:
        header.append(f"件名: {msg.subject}")
    if msg.date:
        header.append(f"日時: {msg.date}")
    if msg.sender:
        header.append(f"差出人: {msg.sender}")
    if msg.to:
        header.append(f"宛先: {msg.to}")
    if msg.cc:
        header.append(f"CC: {msg.cc}")
    raw_body = (msg.body or "").strip()
    # 引用チェーンの重複を除去(新規本文だけ残す)
    body = strip_quoted(raw_body)
    # 添付ファイル名も検索の手がかりとして残す
    attach_names = [a.longFilename or a.shortFilename
                    for a in msg.attachments if (a.longFilename or a.shortFilename)]
    msg.close()
    out = "\n".join(header) + "\n\n---\n\n" + body
    if attach_names:
        out += "\n\n### 添付ファイル\n" + "\n".join(f"- {n}" for n in attach_names)
    return out


def conv_txt(path: Path) -> str:
    for enc in ("utf-8", "cp932", "shift_jis"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(errors="replace")


def conv_html(path: Path) -> str:
    # 文字コードを判定しつつ読み込む
    raw = None
    for enc in ("utf-8", "cp932", "shift_jis"):
        try:
            raw = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raw = path.read_text(errors="replace")

    soup = BeautifulSoup(raw, "html.parser")
    # 本文と無関係な要素を除去(検索ノイズになるため)
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript", "iframe"]):
        tag.decompose()
    # <body> があればその中だけを対象に
    target = soup.body if soup.body else soup
    # リンクは文字だけ残す(URLはノイズになりやすい)、見出しはATX(#)形式
    text = _md(str(target), heading_style="ATX", strip=["a"])
    # 連続する空行を圧縮して整形
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text


# 拡張子 -> 変換関数のディスパッチ表
CONVERTERS = {
    ".docx": conv_docx,
    ".pptx": conv_pptx,
    ".xlsx": conv_xlsx,
    ".pdf":  conv_pdf,
    ".msg":  conv_msg,
    ".txt":  conv_txt,
    ".html": conv_html,
    ".htm":  conv_html,
}


# ---------------------------------------------------------------------------
# メタ情報(YAMLフロントマター)の付与
# ---------------------------------------------------------------------------
def build_frontmatter(src: Path, category: str = "", keywords: str = "") -> str:
    """検索の手がかりになるメタ情報を先頭に付ける。
    category / keywords は後段で Dify に自動生成させる前提で空でも可。"""
    mtime = datetime.datetime.fromtimestamp(src.stat().st_mtime)
    lines = [
        "---",
        f"title: {src.stem}",
        f"source: {src.name}",
        f"date: {mtime:%Y-%m}",
        f"category: {category}",
        f"keywords: {keywords}",
        "---",
        "",
    ]
    return "\n".join(lines)


def safe_name(name: str) -> str:
    """出力ファイル名に使えない文字を除去"""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def source_header(src: Path, part: int = 0, total: int = 1) -> str:
    """本文先頭に置く出典表記。Copilotが根拠ファイルを拾いやすくする。"""
    label = src.name
    if total > 1:
        label += f"（{part}/{total} 分割）"
    return f"> 📄 出典ファイル: **{label}**\n\n"


def source_footer(src: Path) -> str:
    """本文末尾に置く出典表記。回答の根拠として参照されやすくする。"""
    return f"\n\n---\n\n＊この内容の出典ファイル: **{src.name}**\n"


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def process(in_dir: Path, out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)
    targets = [p for p in in_dir.rglob("*")
               if p.is_file() and p.suffix.lower() in CONVERTERS]

    print(f"対象ファイル: {len(targets)} 件\n")
    ok, ng = 0, 0
    for src in sorted(targets):
        try:
            body = CONVERTERS[src.suffix.lower()](src)
            fm = build_frontmatter(src)
            base = safe_name(src.stem)

            # 4MBの壁対策: しきい値を超えたら見出し(##)単位で分割
            if len(fm + body) > SPLIT_THRESHOLD_CHARS:
                chunks = split_by_heading(body, SPLIT_THRESHOLD_CHARS)
                total = len(chunks)
                for idx, chunk in enumerate(chunks, 1):
                    out = out_dir / f"{base}_part{idx}.md"
                    chunk_body = (source_header(src, idx, total)
                                  + chunk
                                  + source_footer(src))
                    out.write_text(fm + chunk_body, encoding="utf-8")
                print(f"  ✓ {src.name}  → {total}分割")
            else:
                full_body = source_header(src) + body + source_footer(src)
                out = out_dir / f"{base}.md"
                out.write_text(fm + full_body, encoding="utf-8")
                print(f"  ✓ {src.name}")
            ok += 1
        except Exception as e:
            print(f"  ✗ {src.name}  ({e})")
            ng += 1

    print(f"\n完了: 成功 {ok} 件 / 失敗 {ng} 件")
    print(f"出力先: {out_dir}")


def split_by_heading(body: str, limit: int) -> list[str]:
    """## 見出し単位でまとめつつ、limit を超えないチャンクに束ねる"""
    parts = re.split(r"(?=^## )", body, flags=re.MULTILINE)
    chunks, cur = [], ""
    for part in parts:
        if len(cur) + len(part) > limit and cur:
            chunks.append(cur)
            cur = part
        else:
            cur += part
    if cur:
        chunks.append(cur)
    return chunks


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使い方: python preprocess_for_copilot.py 入力フォルダ 出力フォルダ")
        sys.exit(1)
    process(Path(sys.argv[1]), Path(sys.argv[2]))
```

======================================================================
## ファイル: enrich_with_dify.py
## 役割: Difyメタ付与＋TLS設定
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dify 連携: Markdown フロントマターへのメタ情報自動付与
======================================================
preprocess_for_copilot.py が出力した Markdown 群を読み込み、
本文を Dify ワークフローに投げて category / keywords / summary を生成し、
フロントマターに書き戻す。

前提となる Dify ワークフロー側の構成:
    入力変数:  text  (文字列。Markdown本文を渡す)
    出力(end): result  (文字列。下記JSONを ensure_ascii=False で文字列化したもの)
        {
          "category": "申請手続き",
          "keywords": ["申請", "承認フロー", "期限"],
          "summary": "申請の承認フローと期限を説明した文書。"
        }
    ※ まさしさんが普段やっている「単一 result キーに json.dumps(..., ensure_ascii=False)」
      の構成をそのまま使える。

使い方:
    # 接続情報は .env に書く(同じフォルダに .env を置く)
    #   DIFY_API_KEY=app-xxxxxxxx
    #   DIFY_BASE_URL=https://api.dify.ai/v1
    #   DIFY_INPUT_VAR=text
    #   DIFY_MAX_CHARS=6000
    #   LOCAL_MARKDOWN_DIR=./output_markdown
    python enrich_with_dify.py                 # .env の LOCAL_MARKDOWN_DIR を使う
    python enrich_with_dify.py  Markdownフォルダ  # 引数で上書きも可
        [--base-url ...] [--input-var ...] [--max-chars N] [--dry-run]
"""

import os
import sys
import re
import json
import time
import argparse
from pathlib import Path

import requests

# .env があれば読み込む(python-dotenv が無くても動くようにフォールバック)
try:
    from dotenv import load_dotenv
    if getattr(sys, "frozen", False):
        _base = Path(sys.executable).resolve().parent
    else:
        _base = Path(__file__).resolve().parent
    load_dotenv(_base / ".env")
except ImportError:
    pass


# ---------------------------------------------------------------------------
# TLS証明書の検証設定(二段構え)
#   1. CA証明書ファイルが指定されていれば、それを使う(ファイル優先)
#   2. 無ければ truststore で OS(Windows)の証明書ストアにフォールバック
#      → IT部門が配布した社内CAなどを Python がそのまま信頼できる
# ---------------------------------------------------------------------------
def setup_tls() -> str | bool:
    """requests の verify に渡す値を返す。
    戻り値: CA証明書ファイルのパス(str) または True(検証する/ストア委譲)。
    環境変数:
      DIFY_CA_BUNDLE       … CA証明書ファイル(.pem/.crt)のパス。あれば最優先。
      DIFY_USE_TRUSTSTORE  … "false" で truststore フォールバックを無効化(既定は有効)
    """
    ca_file = os.getenv("DIFY_CA_BUNDLE", "").strip()
    if ca_file:
        if Path(ca_file).exists():
            print(f"[TLS] 証明書ファイルを使用: {ca_file}")
            return ca_file
        print(f"[TLS] 警告: 指定された証明書ファイルが見つかりません: {ca_file}")
        print("      → OSの証明書ストアにフォールバックします。")

    # ファイル指定が無い/見つからない → truststore で OS ストアに委ねる
    use_ts = os.getenv("DIFY_USE_TRUSTSTORE", "true").strip().lower() != "false"
    if use_ts:
        try:
            import truststore
            truststore.inject_into_ssl()  # 以降 requests は OS の証明書ストアで検証
            print("[TLS] OSの証明書ストア(truststore)を使用します。")
        except ImportError:
            print("[TLS] truststore 未導入。標準のcertifiで検証します。")
            print("      社内CAでエラーが出る場合: pip install truststore")
    return True  # verify=True (truststore注入済みならOSストア、未注入ならcertifi)


# 起動時に一度だけ TLS 設定を解決(injectは早い段階で行うのが望ましい)
_VERIFY = setup_tls()


# ---------------------------------------------------------------------------
# フロントマターの読み書き
# ---------------------------------------------------------------------------
FM_PATTERN = re.compile(r"^---\n(.*?)\n---\n", re.DOTALL)


def split_frontmatter(content: str):
    """(frontmatter_dict, body) を返す。FMが無ければ ({}, content)。"""
    m = FM_PATTERN.match(content)
    if not m:
        return {}, content
    fm_text = m.group(1)
    body = content[m.end():]
    fm = {}
    for line in fm_text.splitlines():
        if ":" in line:
            k, _, v = line.partition(":")
            fm[k.strip()] = v.strip()
    return fm, body


def build_frontmatter(fm: dict) -> str:
    """順序を固定してフロントマターを組み立てる"""
    order = ["title", "source", "date", "category", "keywords", "summary"]
    keys = order + [k for k in fm if k not in order]
    lines = ["---"]
    for k in keys:
        if k in fm:
            lines.append(f"{k}: {fm[k]}")
    lines.append("---")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Dify 呼び出し
# ---------------------------------------------------------------------------
def call_dify(text: str, api_key: str, base_url: str,
              input_var: str = "text", user: str = "enrich-bot",
              retries: int = 3, timeout: int = 120) -> dict:
    """Dify ワークフローを blocking で実行し、result(JSON文字列)を dict にして返す。"""
    url = f"{base_url.rstrip('/')}/workflows/run"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "inputs": {input_var: text},
        "response_mode": "blocking",
        "user": user,
    }

    last_err = None
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=payload,
                                  timeout=timeout, verify=_VERIFY)
            resp.raise_for_status()
            data = resp.json()
            # blocking の戻り: data.outputs.result に格納される構成を想定
            outputs = (data.get("data") or {}).get("outputs") or {}
            raw = outputs.get("result")
            if raw is None:
                # result キーが無い場合、outputs全体から最初の文字列値を拾う保険
                raw = next((v for v in outputs.values() if isinstance(v, str)), "")
            return parse_result(raw)
        except requests.exceptions.RequestException as e:
            last_err = e
            # 429/5xx は待ってリトライ
            wait = 2 ** attempt
            print(f"    リトライ {attempt}/{retries} ({e}) … {wait}s 待機")
            time.sleep(wait)
    raise RuntimeError(f"Dify呼び出し失敗: {last_err}")


def parse_result(raw) -> dict:
    """result を dict にする。文字列ならJSONパース、コードフェンスも除去。"""
    if isinstance(raw, dict):
        return raw
    if not isinstance(raw, str) or not raw.strip():
        return {}
    cleaned = raw.strip()
    # ```json ... ``` を剥がす
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.DOTALL).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        # 末尾の余分なテキストがある場合、最初の {...} だけ抜き出す保険
        m = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    print(f"    ⚠ JSONパース失敗。生の出力: {cleaned[:120]}")
    return {}


# ---------------------------------------------------------------------------
# メタ情報をフロントマターに反映
# ---------------------------------------------------------------------------
def merge_meta(fm: dict, meta: dict) -> dict:
    """Difyが返した category/keywords/summary をフロントマターに反映。
    既に値が入っている項目は上書きしない(手動修正を尊重)。"""
    cat = meta.get("category")
    if cat and not fm.get("category"):
        fm["category"] = str(cat).strip()

    kws = meta.get("keywords")
    if kws and not fm.get("keywords"):
        if isinstance(kws, list):
            fm["keywords"] = ", ".join(str(k).strip() for k in kws)
        else:
            fm["keywords"] = str(kws).strip()

    summary = meta.get("summary")
    if summary and not fm.get("summary"):
        # 改行はフロントマターを壊すので潰す
        fm["summary"] = re.sub(r"\s+", " ", str(summary)).strip()
    return fm


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("md_dir", nargs="?", default=os.getenv("LOCAL_MARKDOWN_DIR"),
                    help="Markdownフォルダ(未指定時は .env の LOCAL_MARKDOWN_DIR)")
    ap.add_argument("--base-url",
                    default=os.getenv("DIFY_BASE_URL", "https://api.dify.ai/v1"))
    ap.add_argument("--input-var",
                    default=os.getenv("DIFY_INPUT_VAR", "text"),
                    help="Difyワークフローの入力変数名")
    ap.add_argument("--max-chars", type=int,
                    default=int(os.getenv("DIFY_MAX_CHARS", "6000")),
                    help="Difyに渡す本文の最大文字数(長文は先頭を要約材料にする)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Difyを呼ばず対象ファイルだけ表示")
    args = ap.parse_args()

    if not args.md_dir:
        print("Markdownフォルダを指定してください(引数 または .env の LOCAL_MARKDOWN_DIR)。")
        sys.exit(1)

    api_key = os.environ.get("DIFY_API_KEY")
    if not api_key and not args.dry_run:
        print("DIFY_API_KEY が未設定です(.env か環境変数で指定してください)。")
        sys.exit(1)

    md_dir = Path(args.md_dir)
    files = sorted(md_dir.glob("*.md"))
    if not files:
        print(f"{md_dir} に .md がありません。")
        sys.exit(1)

    print(f"対象: {len(files)} 件\n")
    done, skip, fail = 0, 0, 0

    for f in files:
        content = f.read_text(encoding="utf-8")
        fm, body = split_frontmatter(content)

        # すでに category も keywords も埋まっていればスキップ(再実行に強い)
        if fm.get("category") and fm.get("keywords"):
            print(f"  - {f.name} (付与済みスキップ)")
            skip += 1
            continue

        if args.dry_run:
            print(f"  ○ {f.name} (dry-run)")
            continue

        # 長文は先頭 max_chars だけ渡す。要約とタグ付けには十分。
        sample = body[:args.max_chars]
        try:
            meta = call_dify(sample, api_key, args.base_url, args.input_var)
            fm = merge_meta(fm, meta)
            new_content = build_frontmatter(fm) + body
            f.write_text(new_content, encoding="utf-8")
            tag = fm.get("category", "?")
            print(f"  ✓ {f.name}  [{tag}]")
            done += 1
        except Exception as e:
            print(f"  ✗ {f.name}  ({e})")
            fail += 1

    print(f"\n完了: 付与 {done} / スキップ {skip} / 失敗 {fail}")


if __name__ == "__main__":
    main()
```

======================================================================
## ファイル: setup.py
## 役割: cx_Freezeビルド設定
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cx_Freeze ビルド設定 — run_pipeline.py を Windows 実行ファイル化する
=====================================================================
Windows 上で、必要ライブラリを入れた環境から:

    pip install cx_Freeze
    python setup.py build

を実行すると、build\\exe.win-amd64-3.xx\\ 以下に
run_pipeline.exe と依存一式が生成される。
そのフォルダごと配布すれば、Python 未導入の PC でも動く。

※ ビルドは「配布先と同じ Windows」で行うこと(cx_Freezeはクロスビルド不可)。
※ ビルド環境には実行に必要なライブラリを全てインストールしておくこと:
   pip install python-docx mammoth python-pptx pymupdf pandas openpyxl \
               extract-msg beautifulsoup4 markdownify requests \
               python-dotenv truststore cx_Freeze
"""

import sys
from cx_Freeze import setup, Executable

# --- 明示的に含めるパッケージ -------------------------------------------
# cx_Freeze は import を辿るが、動的import・データ同梱が絡むものは取りこぼす
# ことがあるため、確実に同梱したいものを packages に列挙する。
packages = [
    # 自作モジュール(run_pipeline がローカルimport している)
    "preprocess_for_copilot",
    "enrich_with_dify",
    # 文書変換系
    "docx",          # python-docx
    "mammoth",
    "pptx",          # python-pptx
    "fitz",          # PyMuPDF(コンパイル済みライブラリを含む)
    "pandas",
    "openpyxl",
    "extract_msg",
    "bs4",           # beautifulsoup4
    "markdownify",
    # 通信・設定・証明書
    "requests",
    "urllib3",
    "certifi",
    "charset_normalizer",
    "idna",
    "dotenv",        # python-dotenv
    "truststore",    # OS証明書ストア検証
    # 標準ライブラリで動的に呼ばれうるもの
    "encodings",
    "ssl",
]

# --- 個別に取りこぼしやすいサブモジュールを includes で補強 -------------
includes = [
    "pandas._libs.tslibs.base",
]

# --- 除外して軽量化(使わない大物) --------------------------------------
excludes = [
    "tkinter",
    "test",
    "unittest",
    "pydoc_data",
    "matplotlib",   # pandasが間接参照する場合があるが今回は不要
]

build_exe_options = {
    "packages": packages,
    "includes": includes,
    "excludes": excludes,
    # 文字化け対策などで zip に含めず展開しておきたいパッケージ
    "zip_include_packages": ["*"],
    "zip_exclude_packages": ["pandas", "fitz", "extract_msg"],
    "include_msvcr": True,   # MSVC ランタイムを同梱(配布先にVC++が無くても動く)
}

# コンソールアプリ(標準出力にログを出すため Win32GUI にはしない)
base = None

setup(
    name="copilot-knowledge-pipeline",
    version="1.0",
    description="案件ナレッジ前処理パイプライン (run_pipeline)",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "run_pipeline.py",
            base=base,
            target_name="run_pipeline.exe",
        )
    ],
)
```

======================================================================
## ファイル: dify_workflow_meta_generator.yml
## 役割: Difyワークフロー定義
======================================================================

```yaml
app:
  name: 業務文書メタ情報ジェネレータ
  description: 業務文書の本文を受け取り、category / keywords / summary を JSON で返す。Copilot向け前処理のメタ付与に使用。
  mode: workflow
  icon: "\U0001F4C4"
  icon_background: "#E0F2FE"
version: "0.1.5"
kind: app

workflow:
  environment_variables: []
  conversation_variables: []
  features:
    file_upload:
      enabled: false
    opening_statement: ""
    retriever_resource:
      enabled: false
    sensitive_word_avoidance:
      enabled: false
    speech_to_text:
      enabled: false
    suggested_questions: []
    suggested_questions_after_answer:
      enabled: false
    text_to_speech:
      enabled: false

  graph:
    edges:
      - id: start-to-llm
        source: start_node
        target: llm_node
        sourceHandle: source
        targetHandle: target
        type: custom
        data:
          sourceType: start
          targetType: llm
          isInLoop: false
      - id: llm-to-end
        source: llm_node
        target: end_node
        sourceHandle: source
        targetHandle: target
        type: custom
        data:
          sourceType: llm
          targetType: end
          isInLoop: false

    nodes:
      # -------- 開始ノード: 入力変数 text --------
      - id: start_node
        type: custom
        position:
          x: 80
          y: 240
        data:
          type: start
          title: 開始
          desc: 業務文書の本文を受け取る
          variables:
            - variable: text
              label: 文書本文
              type: paragraph
              required: true
              max_length: 48000

      # -------- LLMノード: メタ情報をJSONで生成 --------
      - id: llm_node
        type: custom
        position:
          x: 400
          y: 240
        data:
          type: llm
          title: メタ情報生成
          desc: category / keywords / summary を JSON 生成
          model:
            provider: openai
            name: gpt-4o-mini
            mode: chat
            completion_params:
              temperature: 0.2
          prompt_template:
            - role: system
              text: |
                あなたは社内業務文書を分類する専門家です。
                与えられた文書を読み、以下を判定してください。

                - category: 文書が属する業務カテゴリを日本語で1つ（例: 申請手続き, 研修, 顧客対応, 規程, 議事録 など）
                - keywords: 検索の手がかりになる重要語を3〜5個（日本語、配列）
                - summary: 文書の要点を1文で（80文字以内、日本語）

                必ず次のJSON形式のみで出力してください。前置き・後書き・コードブロック記号は一切付けないこと。
                {"category": "...", "keywords": ["...", "..."], "summary": "..."}
            - role: user
              text: |
                以下の業務文書を分類してください。

                {{#start_node.text#}}
          vision:
            enabled: false

      # -------- 終了ノード: result として出力 --------
      - id: end_node
        type: custom
        position:
          x: 720
          y: 240
        data:
          type: end
          title: 終了
          desc: 生成結果を result として返す
          outputs:
            - variable: result
              value_selector:
                - llm_node
                - text
```

======================================================================
## ファイル: projects.example.json
## 役割: 案件設定テンプレート
======================================================================

```json
{
  "_comment": "案件ごとの処理設定。source(元フォルダ)とout(Markdown出力先)はフルパスでOK。場所がバラバラでも指定できる。onedrive_folderはSelenium側がアップする先のOneDrive上のフォルダ名(省略時は案件name)。manifestを省略するとoutフォルダ内に自動作成。",
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

======================================================================
## ファイル: env_sample.txt
## 役割: .envお手本
======================================================================

```text
# =====================================================================
#  設定ファイル（.env のお手本）
# ---------------------------------------------------------------------
#  使い方:
#    1) このファイルを ".env" という名前でコピーする
#         Windows: copy .env.sample .env
#    2) 下の値を自分の環境に合わせて書き換える
#    3) run_pipeline.exe（または run_pipeline.py）と同じフォルダに置く
#
#  ★ .env には API キーなどの秘密情報が入る。共有・Git登録はしないこと。
#  ※ 行頭が # の行はコメント（無視される）。
# =====================================================================


# =====================================================================
#  Dify 接続（メタ情報＝分類・キーワード・要約 の自動付与に使う）
# ---------------------------------------------------------------------
#  Dify を使わない場合は、ここは空のままでよい。
#  その場合は実行時に  --no-dify  を付ける（例: run_pipeline.exe --no-dify）。
# =====================================================================

# Dify アプリの API キー（Difyの「APIアクセス」画面で発行。app- で始まる）
DIFY_API_KEY=app-xxxxxxxxxxxxxxxxxxxx

# Dify の API ベースURL
#   Dify Cloud を使う場合 : https://api.dify.ai/v1
#   社内のセルフホストの場合: https://（社内Difyのホスト名）/v1
DIFY_BASE_URL=https://api.dify.ai/v1

# Dify ワークフローの入力変数名（DSLの開始ノードで定義した名前）
# 既定の dify_workflow_meta_generator.yml を使うなら text のままでよい
DIFY_INPUT_VAR=text

# 1文書あたり Dify に渡す最大文字数（長文は先頭だけ渡す。要約には十分）
DIFY_MAX_CHARS=6000


# =====================================================================
#  TLS 証明書の検証（社内セルフホストDify / 企業プロキシ環境向け）
# ---------------------------------------------------------------------
#  二段構え:
#    1) DIFY_CA_BUNDLE にCA証明書ファイルのパスを指定 → それを最優先で使う
#    2) 指定が無ければ Windows の証明書ストアを使う（社内CAをそのまま信頼）
#  Dify Cloud を普通に使うなら、両方とも下のデフォルトのままでよい。
# =====================================================================

# CA証明書ファイル(.pem/.crt)のパス。通常は空でよい
DIFY_CA_BUNDLE=

# Windowsの証明書ストアを使う（既定 true）。無効にしたいときだけ false
DIFY_USE_TRUSTSTORE=true


# =====================================================================
#  処理の対象・除外の設定
# =====================================================================

# 検索対象から外すフォルダ・ファイルの「接頭辞」（カンマ区切り）
# これで始まるフォルダ名・ファイル名は、配下ごと処理されない。
# 例: _作業中/ や _下書き.docx、▪️メモ/ などを除外
EXCLUDE_PREFIXES=_,▪️,■,~,.


# =====================================================================
#  ファイルの場所（基本は触らなくてよい）
# ---------------------------------------------------------------------
#  どちらも未指定なら、exe（またはスクリプト）と同じフォルダが基準になる。
#  通常は projects.json と upload_list を exe の隣に置くので、空のままでよい。
# =====================================================================

# 案件設定ファイルの場所（既定: exeと同じフォルダの projects.json）
# PROJECTS_FILE=C:/path/to/projects.json

# アップロード対象リストの出力先（既定: exeと同じフォルダの upload_list.json）
# 同じ名前で .md と .csv も同時に出力される
# UPLOAD_LIST_PATH=C:/path/to/upload_list.json


# =====================================================================
#  （補足）LOCAL_MARKDOWN_DIR について
# ---------------------------------------------------------------------
#  これは enrich_with_dify を単体で実行するとき専用の設定。
#  run_pipeline を使う通常運用では不要（案件ごとの出力先は projects.json で
#  指定するため）。単体実行する場合だけ、下のコメントを外して使う。
# =====================================================================
# LOCAL_MARKDOWN_DIR=./output_markdown
```

======================================================================
## ファイル: graph_auth.py
## 役割: 【未使用】Graph認証
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
graph_auth.py — Microsoft Graph 認証＆設定読み込み
===================================================
.env から設定を読み、Graph API 用のアクセストークンを取得する。

2方式に対応:
  - delegated   : 自分の権限。初回だけブラウザ認証、以降は
                  リフレッシュトークンで無人更新(secret不要)。
  - application : アプリ権限。client secret か 証明書で認証
                  (管理者同意が必要)。

機密情報(secret/トークン)はコードに持たず .env と
キャッシュファイルにのみ置く。
"""

import os
import sys
import atexit
from pathlib import Path

from dotenv import load_dotenv
import msal

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
AUTHORITY_TPL = "https://login.microsoftonline.com/{tenant}"

# 委任方式で要求するスコープ。.All なしなので「自分のファイルだけ」。
# offline_access はリフレッシュトークン(無人更新)に必須。
DELEGATED_SCOPES = ["Files.ReadWrite", "User.Read"]
# アプリ方式は固定でこのスコープ(事前に付与した権限が使われる)
APP_SCOPE = ["https://graph.microsoft.com/.default"]


def load_config() -> dict:
    """.env を読み、必須項目を検証して dict で返す。"""
    load_dotenv()
    cfg = {
        "tenant_id": os.getenv("TENANT_ID", "").strip(),
        "client_id": os.getenv("CLIENT_ID", "").strip(),
        "auth_mode": os.getenv("AUTH_MODE", "delegated").strip().lower(),
        "token_cache_path": os.getenv("TOKEN_CACHE_PATH", "./token_cache.bin").strip(),
        "client_secret": os.getenv("CLIENT_SECRET", "").strip(),
        "cert_key_path": os.getenv("CERT_PRIVATE_KEY_PATH", "").strip(),
        "cert_thumbprint": os.getenv("CERT_THUMBPRINT", "").strip(),
        "onedrive_folder": os.getenv("ONEDRIVE_TARGET_FOLDER", "CopilotKnowledge").strip(),
        "sp_hostname": os.getenv("SHAREPOINT_HOSTNAME", "").strip(),
        "sp_site_path": os.getenv("SHAREPOINT_SITE_PATH", "").strip(),
        "sp_folder": os.getenv("SHAREPOINT_TARGET_FOLDER", "CopilotKnowledge").strip(),
        "local_dir": os.getenv("LOCAL_UPLOAD_DIR", "./output_markdown").strip(),
    }

    errors = []
    if not cfg["tenant_id"] or cfg["tenant_id"].startswith("0000"):
        errors.append("TENANT_ID が未設定です。")
    if not cfg["client_id"] or cfg["client_id"].startswith("0000"):
        errors.append("CLIENT_ID が未設定です。")

    if cfg["auth_mode"] == "application":
        has_secret = bool(cfg["client_secret"])
        has_cert = bool(cfg["cert_key_path"] and cfg["cert_thumbprint"])
        if not (has_secret or has_cert):
            errors.append("application方式には CLIENT_SECRET か "
                          "(CERT_PRIVATE_KEY_PATH + CERT_THUMBPRINT) が必要です。")
    elif cfg["auth_mode"] != "delegated":
        errors.append(f"AUTH_MODE は delegated か application。今: {cfg['auth_mode']}")

    if errors:
        print("[設定エラー]")
        for e in errors:
            print("  -", e)
        print("\n.env.example を .env にコピーして値を埋めてください。")
        sys.exit(1)

    return cfg


# ---------------------------------------------------------------------------
# 委任方式: 初回ブラウザ認証 → トークンキャッシュに保存 → 以降無人更新
# ---------------------------------------------------------------------------
def _build_token_cache(path: str) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = Path(path)
    if p.exists():
        cache.deserialize(p.read_text())
    # プロセス終了時、変更があれば書き戻す
    atexit.register(lambda: p.write_text(cache.serialize())
                    if cache.has_state_changed else None)
    return cache


def get_token_delegated(cfg: dict) -> str:
    """委任方式でアクセストークンを取得。
    キャッシュにRTがあれば無人で更新、無ければ初回だけ対話認証。"""
    cache = _build_token_cache(cfg["token_cache_path"])
    app = msal.PublicClientApplication(
        client_id=cfg["client_id"],
        authority=AUTHORITY_TPL.format(tenant=cfg["tenant_id"]),
        token_cache=cache,
    )

    # 1) キャッシュ済みアカウントがあれば silent(無人)で取得を試す
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(DELEGATED_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # 2) 初回 or RT失効 → 対話認証。
    #    まず device flow(ブラウザ別端末でもOK)を試み、ダメなら通常のブラウザ起動。
    print("\n初回認証が必要です。表示される手順でサインインしてください。")
    flow = app.initiate_device_flow(scopes=DELEGATED_SCOPES)
    if "user_code" in flow:
        print("\n" + flow["message"] + "\n")  # URLとコードが表示される
        result = app.acquire_token_by_device_flow(flow)
    else:
        # device flowが無効な環境向けフォールバック
        result = app.acquire_token_interactive(scopes=DELEGATED_SCOPES)

    if "access_token" not in result:
        raise RuntimeError(f"認証失敗: {result.get('error_description', result)}")
    print("認証成功。トークンを保存しました(次回以降は無人で更新されます)。")
    return result["access_token"]


# ---------------------------------------------------------------------------
# アプリ方式: secret か 証明書 でトークン取得(無人)
# ---------------------------------------------------------------------------
def get_token_application(cfg: dict) -> str:
    authority = AUTHORITY_TPL.format(tenant=cfg["tenant_id"])
    if cfg["client_secret"]:
        client_cred = cfg["client_secret"]
    else:
        # 証明書方式: PEM秘密鍵を読み込む
        key_pem = Path(cfg["cert_key_path"]).read_text()
        client_cred = {
            "private_key": key_pem,
            "thumbprint": cfg["cert_thumbprint"],
        }
    app = msal.ConfidentialClientApplication(
        client_id=cfg["client_id"],
        authority=authority,
        client_credential=client_cred,
    )
    result = app.acquire_token_for_client(scopes=APP_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"認証失敗: {result.get('error_description', result)}")
    return result["access_token"]


def get_access_token(cfg: dict) -> str:
    """設定に応じて適切な方式でトークンを返す統一入口。"""
    if cfg["auth_mode"] == "application":
        return get_token_application(cfg)
    return get_token_delegated(cfg)


# 単体テスト: python graph_auth.py で設定検証＆認証確認だけ行う
if __name__ == "__main__":
    cfg = load_config()
    print(f"設定OK (mode={cfg['auth_mode']}, tenant={cfg['tenant_id'][:8]}...)")
    token = get_access_token(cfg)
    print(f"アクセストークン取得OK (長さ {len(token)})")
```

======================================================================
## ファイル: graph_upload.py
## 役割: 【未使用】Graphアップロード
======================================================================

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
graph_upload.py — ローカルフォルダ → OneDrive/SharePoint 自動アップロード
=========================================================================
graph_auth.py で取得したトークンを使い、LOCAL_UPLOAD_DIR 内のファイルを
Graph API 経由で OneDrive(委任) または SharePoint(アプリ) にアップする。

Explorer同期が使えない環境でも、HTTPS(Graph API)経由なので動作する。

使い方:
    python graph_upload.py            # .env の設定に従って全アップロード
    python graph_upload.py --dry-run  # 対象ファイル一覧だけ表示
"""

import sys
import argparse
from pathlib import Path

import requests

from graph_auth import load_config, get_access_token, GRAPH_BASE

# 4MB未満は単純PUT、以上はアップロードセッション(分割)を使う閾値
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
CHUNK_SIZE = 5 * 320 * 1024  # 1.6MB。Graphは320KiBの倍数を要求するため


def get_drive_endpoint(cfg: dict, token: str) -> str:
    """アップロード先ドライブのベースURLを返す。
    委任=自分のOneDrive / アプリ=SharePointサイトのドライブ。"""
    if cfg["auth_mode"] == "application":
        # SharePoint: ホスト名とサイトパスから site を引き、その drive を使う
        if not (cfg["sp_hostname"] and cfg["sp_site_path"]):
            raise RuntimeError("application方式には SHAREPOINT_HOSTNAME と "
                               "SHAREPOINT_SITE_PATH が必要です。")
        site_url = (f"{GRAPH_BASE}/sites/{cfg['sp_hostname']}:"
                    f"/{cfg['sp_site_path'].strip('/')}")
        r = requests.get(site_url, headers=_h(token), timeout=30)
        r.raise_for_status()
        site_id = r.json()["id"]
        return f"{GRAPH_BASE}/sites/{site_id}/drive"
    else:
        # 委任: 自分のOneDrive
        return f"{GRAPH_BASE}/me/drive"


def _h(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


def target_folder(cfg: dict) -> str:
    return (cfg["sp_folder"] if cfg["auth_mode"] == "application"
            else cfg["onedrive_folder"]).strip("/")


def upload_small(drive: str, folder: str, file: Path, token: str):
    """4MB未満: PUT 一発でアップロード。"""
    # パス指定でアップロード。フォルダが無ければGraphが自動作成する。
    url = f"{drive}/root:/{folder}/{file.name}:/content"
    headers = _h(token)
    headers["Content-Type"] = "application/octet-stream"
    r = requests.put(url, headers=headers, data=file.read_bytes(), timeout=120)
    r.raise_for_status()
    return r.json()


def upload_large(drive: str, folder: str, file: Path, token: str):
    """4MB以上: アップロードセッションを作り分割送信。"""
    url = f"{drive}/root:/{folder}/{file.name}:/createUploadSession"
    r = requests.post(url, headers=_h(token), json={
        "item": {"@microsoft.graph.conflictBehavior": "replace"}
    }, timeout=30)
    r.raise_for_status()
    upload_url = r.json()["uploadUrl"]

    size = file.stat().st_size
    with open(file, "rb") as f:
        start = 0
        while start < size:
            chunk = f.read(CHUNK_SIZE)
            end = start + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{size}",
            }
            cr = requests.put(upload_url, headers=headers, data=chunk, timeout=120)
            # 202=継続, 200/201=完了
            if cr.status_code not in (200, 201, 202):
                cr.raise_for_status()
            start = end + 1
    return {"name": file.name}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    cfg = load_config()
    local = Path(cfg["local_dir"])
    if not local.exists():
        print(f"アップロード元フォルダがありません: {local}")
        sys.exit(1)

    files = [p for p in sorted(local.glob("*")) if p.is_file()]
    if not files:
        print(f"{local} にファイルがありません。")
        sys.exit(0)

    folder = target_folder(cfg)
    print(f"アップロード元 : {local}")
    print(f"アップロード先 : {cfg['auth_mode']} / フォルダ「{folder}」")
    print(f"対象ファイル   : {len(files)} 件\n")

    if args.dry_run:
        for f in files:
            mb = f.stat().st_size / 1024 / 1024
            print(f"  ○ {f.name} ({mb:.2f}MB)")
        print("\n(dry-run: 実際のアップロードは行っていません)")
        return

    token = get_access_token(cfg)
    drive = get_drive_endpoint(cfg, token)

    ok, ng = 0, 0
    for f in files:
        try:
            if f.stat().st_size < SIMPLE_UPLOAD_LIMIT:
                upload_small(drive, folder, f, token)
            else:
                upload_large(drive, folder, f, token)
            print(f"  ✓ {f.name}")
            ok += 1
        except requests.exceptions.HTTPError as e:
            body = e.response.text[:200] if e.response is not None else ""
            print(f"  ✗ {f.name}  HTTP {e.response.status_code if e.response else '?'}: {body}")
            ng += 1
        except Exception as e:
            print(f"  ✗ {f.name}  ({e})")
            ng += 1

    print(f"\n完了: 成功 {ok} / 失敗 {ng}")


if __name__ == "__main__":
    main()
```