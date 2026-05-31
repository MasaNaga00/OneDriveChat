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
