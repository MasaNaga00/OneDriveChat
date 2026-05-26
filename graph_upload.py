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
