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
