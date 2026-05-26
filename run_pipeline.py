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

# .env 読み込み
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# 既存スクリプトの部品を再利用(車輪の再発明をしない)
import preprocess_for_copilot as pp
from enrich_with_dify import (
    split_frontmatter, build_frontmatter, call_dify, merge_meta,
)

MANIFEST_VERSION = 1


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
              "deleted": 0, "enriched": 0, "error": None}

    print(f"\n==== 案件: {name} ====")
    if not source_dir.exists():
        print(f"  ⚠ 元フォルダがありません: {source_dir}  → スキップ")
        result["error"] = "source not found"
        return result
    out_dir.mkdir(parents=True, exist_ok=True)

    use_dify = not opts["no_dify"]
    src_files = [p for p in sorted(source_dir.rglob("*"))
                 if p.is_file() and p.suffix.lower() in pp.CONVERTERS]

    manifest = {} if opts["force"] else load_manifest(manifest_path)
    new_manifest = {}
    seen_keys = set()

    print(f"  元: {source_dir}")
    print(f"  出力: {out_dir}")
    print(f"  対象 {len(src_files)} 件 (台帳 {len(manifest)} 件)")

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
    ap.add_argument("--projects", default=os.getenv("PROJECTS_FILE", "./projects.json"),
                    help="案件設定JSON(既定 ./projects.json または .env の PROJECTS_FILE)")
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


if __name__ == "__main__":
    main()
