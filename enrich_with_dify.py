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
    export DIFY_API_KEY="app-xxxxxxxx"
    python enrich_with_dify.py  Markdownフォルダ
        [--base-url https://api.dify.ai/v1]
        [--max-chars 6000]
        [--dry-run]
"""

import os
import sys
import re
import json
import time
import argparse
from pathlib import Path

import requests


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
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
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
    ap.add_argument("md_dir", help="Markdownフォルダ(preprocess_for_copilotの出力)")
    ap.add_argument("--base-url", default="https://api.dify.ai/v1")
    ap.add_argument("--input-var", default="text", help="Difyワークフローの入力変数名")
    ap.add_argument("--max-chars", type=int, default=6000,
                    help="Difyに渡す本文の最大文字数(長文は先頭を要約材料にする)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Difyを呼ばず対象ファイルだけ表示")
    args = ap.parse_args()

    api_key = os.environ.get("DIFY_API_KEY")
    if not api_key and not args.dry_run:
        print("環境変数 DIFY_API_KEY が未設定です。")
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
