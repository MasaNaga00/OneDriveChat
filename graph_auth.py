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
