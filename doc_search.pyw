# -*- coding: utf-8 -*-
# doc_search.py — 文書内検索 + 2ペイン・タブ式エクスプローラ（標準ライブラリのみ）
# 実行: python doc_search.py             画面を起動 (.pyw にリネームして運用可)
#       python doc_search.py --scan      フォルダインデックスだけ更新(夜間バッチ用)
#       python doc_search.py --make-icon exe化用の doc_search.ico を書き出す
# 視覚: 種類別アイコン(フォルダ/Excel/Word/PowerPoint/PDF/テキスト/コード/ZIP)を
#       画像ファイルなしでコード描画。「種類」列もExplorer同様の型名表示
# ---- キーボード操作(Windowsエクスプローラ準拠) ----
#   F6 / Shift+F6 : ペイン切替(アドレス→検索→左ペイン→一覧)
#   Ctrl+L / Alt+D: アドレスバーへ
#   Ctrl+F: フォルダ名検索欄   Ctrl+E / F3: ファイル名検索欄
#   Ctrl+B: お気に入り絞り込み(左ペイン)
#   Ctrl+O / Ctrl+K: お気に入りペインへ / ファイル一覧へ
#   Alt+← / →     : 戻る / 進む     BackSpace: 戻る   Alt+↑: 一つ上へ
#   Alt+Home      : ホーム           Ctrl+R: 再読み込み
#   Ctrl+T: 新しいタブ  Ctrl+W: タブを閉じる
#   Ctrl+Tab / Ctrl+Shift+Tab: 次/前のタブ  Ctrl+1〜9: 割り当てたタブへ
#   Ctrl+Shift+D: 2ペイン表示切替  Tab(一覧上): 反対のペインへ
#   Ctrl+X / Ctrl+C / Ctrl+V / Delete (一覧上): 切取/コピー/貼付/ごみ箱
#   F2(一覧上): 名前の変更   Ctrl+Shift+N: 新しいフォルダー
#   Ctrl+A(一覧上): すべて選択   Ctrl+H: ~$一時ファイルの表示/非表示
#   Explorer・デスクトップからのドラッグ&ドロップ受け入れ(コピー)
#   一覧で Enter  : フォルダに入る / ファイルを開く
#   Ctrl+Enter    : Explorerで場所を開く
#   アプリケーションキー / Shift+F10: 右クリックメニュー
#   F2:文書内検索フォルダ欄  F4:除外欄  F5:検索実行  F7:お気に入り
#   F8:選択項目をお気に入り追加  F9:CSV出力  Ctrl+D:参照  F1:一覧
#   Ctrl+Shift+K/O/I/P: 拡張子保存/お気に入り保存/インデックス作成/スキャン停止
#   Alt+1/2/3/4: 文書内検索/エクスプローラ/フォルダ検索/設定タブへ
#   Ctrl+G: この場所で文書内検索
#   Ctrl+Shift+E: サクラエディタで開く  Ctrl+Shift+C / S: ここでcmd / PowerShell
#   Alt+Enter: プロパティ  Shift+Delete: 完全削除  一覧で文字入力: 頭文字ジャンプ
#   アドレスバーに cmd / powershell / wt と入力+Enter でもその場所で端末を開く
# ---- フォルダ検索タブ(Everything風) ----
#   入力するそばからインデックス全体のフォルダ名/ファイル名を検索して一覧表示
#   Enter/ダブルクリック: フォルダはエクスプローラタブで、ファイルはそのまま開く
#   Ctrl+Enter: Explorerで開く  Ctrl+Shift+F: 対象(フォルダ/ファイル/両方)切替
# ---- 検索語の書き方(フォルダ名/ファイル名検索共通) ----
#   スペース区切りで AND、-語 で除外、* ? を含む語はワイルドカード一致
#   フォルダ名検索は既定で「フォルダ名そのもの」に一致(パス全体は「パス」にチェック)
#   \ や / を含む語は常にパス全体に対して照合する
import os
import re
import csv
import sys
import html
import time
import heapq
import queue
import bisect
import shutil
import zipfile
import functools
import unicodedata
import threading
import subprocess
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog

CTX = 20
MAX_HIT = 20
LIST_DEPTH = 3
MAX_ROWS = 2000
MAX_LIST = 200000
SCAN_WORKERS = 8
FIND_DELAY_MS = 150      # フォルダ検索タブ: 入力が止まってから検索するまでの待ち
# フォルダ検索タブの「対象」
FIND_MODES = {"dir": "フォルダ", "file": "ファイル", "both": "両方"}
# 上位タブの並び順(Alt+1〜4 もこの順)
TAB_DOC, TAB_EXP, TAB_FIND, TAB_CONF = range(4)
# インデックス作成時に潜らないシステムフォルダ(小文字)
SKIP_DIRS = {"$recycle.bin", "system volume information", "$windows.~bt",
             "$windows.~ws", "windows.old"}
TEXT_EXT = (".txt", ".md", ".csv", ".tsv", ".log", ".ini", ".conf",
            ".py", ".sql", ".xml", ".json", ".yaml", ".yml",
            ".html", ".htm", ".bat", ".ps1", ".ttl", ".java", ".c", ".js")
# exe化(PyInstaller)時は展開先ではなくexeの置き場所を基準にする
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXT_FILE = os.path.join(BASE_DIR, "doc_search_ext.txt")
CONF_FILE = os.path.join(BASE_DIR, "doc_search_conf.txt")
FAV_FILE = os.path.join(BASE_DIR, "doc_search_fav.txt")
ROOTS_FILE = os.path.join(BASE_DIR, "doc_search_roots.txt")
DIRIDX_FILE = os.path.join(BASE_DIR, "doc_search_dirindex.txt")
FILEIDX_FILE = os.path.join(BASE_DIR, "doc_search_fileindex.txt")
TABS_FILE = os.path.join(BASE_DIR, "doc_search_tabs.txt")

WIN_STD = "Windows標準"
THEMES = {
    WIN_STD: {
        "dark": False,
        "bg": "SystemButtonFace", "panel": "SystemButtonFace",
        "field": "SystemWindow", "alt": "SystemWindow",
        "fg": "SystemWindowText", "sub": "#6d6d6d",
        "accent": "SystemHighlight", "accent_dk": "#adadad",
        "sel": "SystemHighlight", "selfg": "SystemHighlightText",
        "folder": "SystemWindowText", "insert": "SystemWindowText",
        "tabsel": "SystemWindowText", "status": "SystemWindowText",
        "gold": "#d9d9d9",
    },
    "ダーク（シンプル）": {
        "dark": True,
        "bg": "#252526", "panel": "#2d2d30", "field": "#1f1f1f",
        "alt": "#1f1f1f", "fg": "#e6e6e6", "sub": "#9b9b9b",
        "accent": "#5a9fd4", "accent_dk": "#3c6f96",
        "sel": "#264f78", "selfg": "#ffffff",
        "folder": "#e6e6e6", "insert": "#d4d4d4",
        "tabsel": "#ffffff", "status": "#bfbfbf",
        "gold": "#3c3c3c",
    },
    "チャコール×水色": {
        "dark": True,
        "bg": "#2a2e35", "panel": "#22262d", "field": "#1a1d22",
        "alt": "#20242a", "fg": "#e9ecef", "sub": "#8f9aa6",
        "accent": "#63cdf6", "accent_dk": "#2f89ad",
        "sel": "#17455c", "selfg": "#63cdf6",
        "folder": "#63cdf6", "insert": "#63cdf6",
        "tabsel": "#63cdf6", "status": "#63cdf6",
        "gold": "#d9b74a",
    },
    "黒×赤": {
        "dark": True,
        "bg": "#151517", "panel": "#0e0e10", "field": "#1d1d21",
        "alt": "#19191d", "fg": "#ededee", "sub": "#87878e",
        "accent": "#ff4757", "accent_dk": "#b3202c",
        "sel": "#4a1219", "selfg": "#ff4757",
        "folder": "#ff4757", "insert": "#ff4757",
        "tabsel": "#ff4757", "status": "#ff4757",
        "gold": "#d9b74a",
    },
}
SYS_FALL = {"SystemButtonFace": "#f0f0f0", "SystemWindow": "#ffffff",
            "SystemWindowText": "#000000", "SystemHighlight": "#0078d7",
            "SystemHighlightText": "#ffffff"}

JP_KEYS = ("yu gothic", "yu mincho", "meiryo", "ms gothic", "ms pgothic",
           "ms mincho", "ms pmincho", "ms ui gothic", "biz ud",
           "noto sans jp", "noto serif jp", "noto sans cjk",
           "noto serif cjk", "ipa", "takao", "vl gothic", "migu",
           "ricty", "sarasa", "udev gothic", "hackgen", "plemol")
DEFAULT_FONTS = ("Yu Gothic UI", "游ゴシック", "BIZ UDPゴシック",
                 "BIZ UDPGothic", "メイリオ", "Meiryo", "MS UI Gothic")

HELP_TEXT = """【全体】
  F1                 このショートカット一覧
  Alt+1 / 2 / 3 / 4  文書内検索 / エクスプローラ / フォルダ検索 / 設定タブへ
  Ctrl+G             今いる場所を対象に文書内検索へ
  F5                 検索実行(表示中のタブに応じて)
  F6 / Shift+F6      ペイン切替(アドレス→検索→左ペイン→一覧)

【文書内検索】
  F2 / F3 / F4       フォルダ欄 / 検索語欄 / 除外欄へ
  Ctrl+D             参照ダイアログ
  F7 / F8            お気に入りメニュー / お気に入りに追加
  F9                 CSV出力
  Esc                検索を停止(検索中)
  一覧: ダブルクリック / Ctrl+Enter=ファイルを開く  Enter=フォルダを開く
  Ctrl+Shift+E       ヒットしたファイルをサクラエディタで開く
  ※ 大文字小文字は区別しません。「全角/半角を区別しない」で ＡＢＣ と ABC、
     ｱｲｳ と アイウ も同じとみなします。Excelは数値セルも検索対象です

【フォルダ検索】(Everything風: 入力するそばから全フォルダ/ファイルを検索)
  対象               フォルダ / ファイル / 両方 (Ctrl+Shift+F で切替)
                     ファイルは設定タブ「ファイル名もインデックスに含める」が必要
  Ctrl+F / F3        検索欄へ     ↓ (検索欄)  結果一覧へ    Esc  検索欄を空に
  Enter / ダブルクリック  フォルダ: エクスプローラタブで開く
                         ファイル: そのファイルを開く
  Shift+Enter        新しいタブで開く    Ctrl+Enter  Explorerで開く(選択)
  Ctrl+C             フルパスをコピー    Ctrl+R      再検索
  列見出しクリック   名前/場所で並べ替え(3回目で関連度順に戻る)
  ※ 対象は設定タブの「検索ルート」から作ったインデックス(Ctrl+Shift+I で更新)

【エクスプローラ】
  Ctrl+L / Alt+D     アドレスバーへ(cmd / powershell / wt と入力で端末起動)
  Ctrl+F             フォルダ名検索欄へ(全フォルダから場所探し)
  Ctrl+E / F3        ファイル名検索欄へ(今の場所から深さN層。空欄なら通常表示)
  検索語の書き方     スペース区切り=AND  -語=除外  *や?=ワイルドカード
                     フォルダ名検索は「パス」にチェックでパス全体を対象
                     (\\ や / を含む語は常にパス全体で照合)
  BackSpace          フォルダ名検索の結果から開いた場所なら結果一覧へ戻る
  Ctrl+Shift+E       サクラエディタで開く
  Ctrl+Shift+C / S   この場所でコマンドプロンプト / PowerShell
  Alt+Enter          プロパティ     Shift+Delete  完全削除(ごみ箱を経由しない)
  文字キー(一覧上)   頭文字ジャンプ(続けて打つと絞り込み)
  Ctrl+B             お気に入り絞り込み(左ペイン)
  Ctrl+O / Ctrl+K    お気に入りペインへ / ファイル一覧へ
  F8                 選択中の行(またはこの場所)をお気に入りに追加
  左ペイン: ドラッグで並べ替え/グループ移動  Ctrl+↑↓でも移動可
  Alt+← / → / ↑      戻る / 進む / 一つ上へ
  BackSpace          戻る        Alt+Home  ホーム
  Ctrl+R             再読み込み
  Ctrl+Shift+D       2ペイン表示の切替
  Tab (一覧上)       反対のペインへ
  Ctrl+X / Ctrl+C    切り取り(貼り付けで移動) / コピー
  Ctrl+V             このフォルダへ貼り付け
  Delete             ごみ箱へ削除
  F2 (一覧上)        名前の変更
  Ctrl+Shift+N       新しいフォルダー
  Ctrl+H             ~$で始まる一時ファイル(Excel等の編集中)の表示/非表示
  Ctrl+A (一覧上)    すべて選択(Shift/Ctrl+クリックで個別複数選択)
  ※ Explorerやデスクトップからのドラッグ&ドロップでコピーできます
  一覧: Enter=開く  Ctrl+Enter=Explorerで場所を開く
  アプリキー / Shift+F10  右クリックメニュー

【タブ】※どの画面からでも有効(アクティブなペインに作用)
  Ctrl+T / Ctrl+W    新しいタブ / タブを閉じる
  Ctrl+Tab / +Shift  次のタブ / 前のタブ
  Ctrl+1〜9          割り当てたタブへ(タブ右クリックで設定)

【設定】
  Ctrl+Shift+K       拡張子を保存
  Ctrl+Shift+O       お気に入りを保存
  Ctrl+Shift+I       インデックス作成/更新
  Ctrl+Shift+P       スキャン停止
"""

# 拡張子 -> (アイコン種別, 種類列の表示名)
TYPE_MAP = {
    ".xlsx": ("excel", "Excel"), ".xlsm": ("excel", "Excel"),
    ".xls": ("excel", "Excel"), ".csv": ("excel", "CSV"),
    ".tsv": ("excel", "TSV"),
    ".docx": ("word", "Word"), ".docm": ("word", "Word"),
    ".doc": ("word", "Word"),
    ".pptx": ("ppt", "PowerPoint"), ".pptm": ("ppt", "PowerPoint"),
    ".ppt": ("ppt", "PowerPoint"),
    ".pdf": ("pdf", "PDF"),
    ".zip": ("zip", "ZIP"), ".7z": ("zip", "7-Zip"), ".lzh": ("zip", "LZH"),
    ".txt": ("text", "テキスト"), ".md": ("text", "Markdown"),
    ".log": ("text", "ログ"), ".ini": ("text", "INI"),
    ".conf": ("text", "CONF"), ".ttl": ("text", "TTLマクロ"),
    ".py": ("code", "Python"), ".sql": ("code", "SQL"),
    ".xml": ("code", "XML"), ".json": ("code", "JSON"),
    ".yaml": ("code", "YAML"), ".yml": ("code", "YAML"),
    ".html": ("code", "HTML"), ".htm": ("code", "HTML"),
    ".bat": ("code", "バッチ"), ".ps1": ("code", "PowerShell"),
    ".java": ("code", "Java"), ".c": ("code", "C"),
    ".js": ("code", "JavaScript"),
}


def file_type(name):
    """ファイル名 -> (アイコン種別, 種類表示名)"""
    ext = os.path.splitext(name)[1].lower()
    if ext in TYPE_MAP:
        return TYPE_MAP[ext]
    if ext:
        return ("file", ext[1:].upper())
    return ("file", "ファイル")


# ============================================================
#  アイコン(画像ファイル不要・コードで描画)
# ============================================================
def _rect(img, x1, y1, x2, y2, c):
    for y in range(y1, y2 + 1):
        for x in range(x1, x2 + 1):
            img.put(c, (x, y))


def _sheet(img):
    """書類の下地(白い紙+グレー枠+折り目)"""
    _rect(img, 3, 1, 12, 14, "#ffffff")
    _rect(img, 3, 1, 12, 1, "#9a9a9a")
    _rect(img, 3, 14, 12, 14, "#9a9a9a")
    _rect(img, 3, 1, 3, 14, "#9a9a9a")
    _rect(img, 12, 1, 12, 14, "#9a9a9a")
    _rect(img, 10, 2, 11, 3, "#d9d9d9")


def _sheet_with(band=None):
    im = tk.PhotoImage(width=16, height=16)
    _sheet(im)
    if band:
        _rect(im, 1, 8, 8, 14, band)
    return im


def build_icons():
    ic = {}

    im = tk.PhotoImage(width=16, height=16)      # フォルダ
    _rect(im, 2, 3, 7, 5, "#d8a136")
    _rect(im, 2, 5, 13, 13, "#f2c14e")
    _rect(im, 2, 5, 13, 5, "#e0ac37")
    _rect(im, 2, 13, 13, 13, "#c9962a")
    _rect(im, 2, 5, 2, 13, "#c9962a")
    _rect(im, 13, 5, 13, 13, "#c9962a")
    ic["folder"] = im

    ic["file"] = _sheet_with()                   # 汎用ファイル

    im = _sheet_with()                           # テキスト(罫線)
    for y in (4, 7, 10):
        _rect(im, 5, y, 10, y, "#8a8a8a")
    ic["text"] = im

    ic["excel"] = _sheet_with("#217346")         # Excel(緑)
    ic["word"] = _sheet_with("#2b579a")          # Word(青)
    ic["ppt"] = _sheet_with("#d24726")           # PowerPoint(橙)
    ic["pdf"] = _sheet_with("#c11e1e")           # PDF(赤)

    im = _sheet_with("#3f4a5a")                  # コード(<>印)
    for x, y in ((4, 10), (3, 11), (4, 12), (6, 10), (7, 11), (6, 12)):
        im.put("#ffffff", (x, y))
    ic["code"] = im

    im = tk.PhotoImage(width=16, height=16)      # ZIP(箱+ジッパー)
    _rect(im, 3, 2, 12, 13, "#caa55a")
    _rect(im, 3, 2, 12, 2, "#a67f35")
    _rect(im, 3, 13, 12, 13, "#a67f35")
    _rect(im, 3, 2, 3, 13, "#a67f35")
    _rect(im, 12, 2, 12, 13, "#a67f35")
    for y in range(3, 12, 2):
        im.put("#6b5220", (7, y))
        im.put("#6b5220", (8, y + 1))
    ic["zip"] = im
    return ic


# ============================================================
#  アプリ自体のアイコン(タイトルバー/タスクバー/exe用ico)
# ============================================================
def app_icon_rows(s):
    """サイズsのアプリアイコンを色マトリクスで返す(None=透過)"""
    G = [[None] * s for _ in range(s)]

    def rect(x1, y1, x2, y2, c):
        for y in range(max(0, y1), min(s, y2 + 1)):
            for x in range(max(0, x1), min(s, x2 + 1)):
                G[y][x] = c

    def ring(cx, cy, ro, ri, c):
        for y in range(s):
            for x in range(s):
                d = (x - cx) ** 2 + (y - cy) ** 2
                if ri * ri <= d <= ro * ro:
                    G[y][x] = c

    u = s / 16.0
    b = max(1, int(u / 2))
    rect(int(1 * u), int(3 * u), int(7 * u), int(5 * u), "#d8a136")
    rect(int(1 * u), int(5 * u), int(15 * u) - 1, int(13 * u), "#f2c14e")
    rect(int(1 * u), int(5 * u), int(15 * u) - 1, int(5 * u) + b, "#e0ac37")
    rect(int(1 * u), int(13 * u) - b, int(15 * u) - 1, int(13 * u),
         "#c9962a")
    rect(int(1 * u), int(5 * u), int(1 * u) + b - 1, int(13 * u), "#c9962a")
    rect(int(15 * u) - b, int(5 * u), int(15 * u) - 1, int(13 * u),
         "#c9962a")
    cx, cy = int(10.3 * u), int(9.0 * u)
    ring(cx, cy, 3.3 * u, 2.1 * u, "#2f6b9e")
    w = max(1, int(0.9 * u))
    step = int(2.4 * u)
    for i in range(step):
        x = int(cx + 2.3 * u) + i
        y = int(cy + 2.3 * u) + i
        rect(x, y, x + w, y + w, "#2f6b9e")
    return G


def rows_to_photo(rows):
    s = len(rows)
    img = tk.PhotoImage(width=s, height=s)
    for y, row in enumerate(rows):
        for x, c in enumerate(row):
            if c:
                img.put(c, (x, y))
    return img


def write_ico(path, sizes=(16, 32, 48)):
    """外部ライブラリなしで .ico を書き出す(32bit BGRA)"""
    import struct
    images = []
    for s in sizes:
        rows = app_icon_rows(s)
        px = bytearray()
        for y in range(s - 1, -1, -1):
            for x in range(s):
                c = rows[y][x]
                if c:
                    r = int(c[1:3], 16)
                    g = int(c[3:5], 16)
                    bch = int(c[5:7], 16)
                    px += bytes((bch, g, r, 255))
                else:
                    px += bytes((0, 0, 0, 0))
        stride = ((s + 31) // 32) * 4
        mask = bytes(stride * s)
        header = struct.pack("<IiiHHIIiiII", 40, s, s * 2, 1, 32, 0,
                             len(px) + len(mask), 0, 0, 0, 0)
        images.append(header + bytes(px) + mask)
    out = struct.pack("<HHH", 0, 1, len(sizes))
    offset = 6 + 16 * len(sizes)
    entries = b""
    for s, data in zip(sizes, images):
        entries += struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0,
                               1, 32, len(data), offset)
        offset += len(data)
    with open(path, "wb") as f:
        f.write(out + entries)
        for d in images:
            f.write(d)


def set_app_id():
    """タスクバーでPythonではなく本アプリとして扱わせる(Windows)"""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Base.DocSearch.Explorer")
    except Exception:
        pass


# ============================================================
#  Windowsファイル操作(クリップボード/ごみ箱/ドロップ受け入れ)
# ============================================================
def clip_set_files(paths, move=False):
    """ファイル群をCF_HDROPでクリップボードへ(Explorerに貼り付け可)
    move=True で「切り取り」(貼り付け側が移動として扱う)"""
    try:
        import ctypes
        CF_HDROP = 15
        GMEM_MOVEABLE = 0x2
        files = "\0".join(paths) + "\0\0"
        data = files.encode("utf-16-le")
        header = (20).to_bytes(4, "little") + bytes(12) \
            + (1).to_bytes(4, "little")
        payload = header + data
        k32 = ctypes.windll.kernel32
        u32 = ctypes.windll.user32
        k32.GlobalAlloc.restype = ctypes.c_void_p
        k32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        k32.GlobalLock.restype = ctypes.c_void_p
        k32.GlobalLock.argtypes = [ctypes.c_void_p]
        k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        u32.SetClipboardData.restype = ctypes.c_void_p
        u32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        u32.RegisterClipboardFormatW.restype = ctypes.c_uint
        u32.RegisterClipboardFormatW.argtypes = [ctypes.c_wchar_p]
        h = k32.GlobalAlloc(GMEM_MOVEABLE, len(payload))
        p = k32.GlobalLock(h)
        ctypes.memmove(p, payload, len(payload))
        k32.GlobalUnlock(h)
        eff = (2 if move else 5).to_bytes(4, "little")
        h2 = k32.GlobalAlloc(GMEM_MOVEABLE, 4)
        p2 = k32.GlobalLock(h2)
        ctypes.memmove(p2, eff, 4)
        k32.GlobalUnlock(h2)
        if not u32.OpenClipboard(0):
            return False
        u32.EmptyClipboard()
        u32.SetClipboardData(CF_HDROP, h)
        fmt = u32.RegisterClipboardFormatW("Preferred DropEffect")
        if fmt:
            u32.SetClipboardData(fmt, h2)
        u32.CloseClipboard()
        return True
    except Exception:
        return False


def clip_get_files():
    """クリップボードのファイル一覧を取得 -> (パスのlist, 移動フラグ)"""
    try:
        import ctypes
        CF_HDROP = 15
        u32 = ctypes.windll.user32
        s32 = ctypes.windll.shell32
        k32 = ctypes.windll.kernel32
        u32.GetClipboardData.restype = ctypes.c_void_p
        k32.GlobalLock.restype = ctypes.c_void_p
        k32.GlobalLock.argtypes = [ctypes.c_void_p]
        k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        s32.DragQueryFileW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                       ctypes.c_wchar_p, ctypes.c_uint]
        u32.RegisterClipboardFormatW.restype = ctypes.c_uint
        u32.RegisterClipboardFormatW.argtypes = [ctypes.c_wchar_p]
        if not u32.IsClipboardFormatAvailable(CF_HDROP):
            return [], False
        if not u32.OpenClipboard(0):
            return [], False
        try:
            h = u32.GetClipboardData(CF_HDROP)
            if not h:
                return [], False
            n = s32.DragQueryFileW(h, 0xFFFFFFFF, None, 0)
            out = []
            for i in range(n):
                ln = s32.DragQueryFileW(h, i, None, 0)
                buf = ctypes.create_unicode_buffer(ln + 1)
                s32.DragQueryFileW(h, i, buf, ln + 1)
                out.append(buf.value)
            move = False
            fmt = u32.RegisterClipboardFormatW("Preferred DropEffect")
            if fmt and u32.IsClipboardFormatAvailable(fmt):
                h2 = u32.GetClipboardData(fmt)
                if h2:
                    pp = k32.GlobalLock(h2)
                    if pp:
                        val = ctypes.cast(
                            pp, ctypes.POINTER(ctypes.c_uint)
                        ).contents.value
                        k32.GlobalUnlock(h2)
                        move = (val & 2) != 0 and (val & 1) == 0
            return out, move
        finally:
            u32.CloseClipboard()
    except Exception:
        return [], False


def clip_clear():
    """クリップボードを空にする(切り取り→移動完了後に使用)"""
    try:
        import ctypes
        u32 = ctypes.windll.user32
        if u32.OpenClipboard(0):
            u32.EmptyClipboard()
            u32.CloseClipboard()
    except Exception:
        pass


def delete_to_trash(paths, permanent=False):
    """ごみ箱へ削除(Windows標準の確認ダイアログ付き)
    permanent=True でごみ箱を経由せず完全削除(Shift+Delete相当)"""
    try:
        import ctypes

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [("hwnd", ctypes.c_void_p),
                        ("wFunc", ctypes.c_uint),
                        ("pFrom", ctypes.c_wchar_p),
                        ("pTo", ctypes.c_wchar_p),
                        ("fFlags", ctypes.c_ushort),
                        ("fAnyOperationsAborted", ctypes.c_int),
                        ("hNameMappings", ctypes.c_void_p),
                        ("lpszProgressTitle", ctypes.c_wchar_p)]
        FO_DELETE = 3
        FOF_ALLOWUNDO = 0x40
        op = SHFILEOPSTRUCTW()
        op.hwnd = None
        op.wFunc = FO_DELETE
        op.pFrom = "\0".join(paths) + "\0"
        op.pTo = None
        op.fFlags = 0 if permanent else FOF_ALLOWUNDO
        res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        return res == 0 and not op.fAnyOperationsAborted
    except Exception:
        return False


def shell_verb(path, verb):
    """Explorerのシェル動詞を実行。verb="properties"でプロパティ、
    "openas"で「プログラムから開く」ダイアログ"""
    try:
        import ctypes

        class SHELLEXECUTEINFOW(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_ulong),
                        ("fMask", ctypes.c_ulong),
                        ("hwnd", ctypes.c_void_p),
                        ("lpVerb", ctypes.c_wchar_p),
                        ("lpFile", ctypes.c_wchar_p),
                        ("lpParameters", ctypes.c_wchar_p),
                        ("lpDirectory", ctypes.c_wchar_p),
                        ("nShow", ctypes.c_int),
                        ("hInstApp", ctypes.c_void_p),
                        ("lpIDList", ctypes.c_void_p),
                        ("lpClass", ctypes.c_wchar_p),
                        ("hkeyClass", ctypes.c_void_p),
                        ("dwHotKey", ctypes.c_ulong),
                        ("hIcon", ctypes.c_void_p),
                        ("hProcess", ctypes.c_void_p)]
        SEE_MASK_INVOKEIDLIST = 0x0000000C
        SW_SHOWNORMAL = 1
        info = SHELLEXECUTEINFOW()
        info.cbSize = ctypes.sizeof(info)
        info.fMask = SEE_MASK_INVOKEIDLIST
        info.hwnd = None
        info.lpVerb = verb
        info.lpFile = path
        info.lpParameters = None
        info.lpDirectory = None
        info.nShow = SW_SHOWNORMAL
        return bool(ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info)))
    except Exception:
        return False


def list_drives():
    """接続中のドライブ -> [(表示名, "C:\\"), ...]"""
    out = []
    if os.name != "nt":
        return out
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        mask = k32.GetLogicalDrives()
        for i in range(26):
            if not mask & (1 << i):
                continue
            root = "%s:\\" % chr(65 + i)
            label = ""
            try:
                buf = ctypes.create_unicode_buffer(261)
                if k32.GetVolumeInformationW(root, buf, 261, None, None,
                                             None, None, 0):
                    label = buf.value
            except Exception:
                pass
            dtype = k32.GetDriveTypeW(root)
            if not label:
                label = {2: "リムーバブル ディスク", 3: "ローカル ディスク",
                         4: "ネットワーク ドライブ", 5: "CD ドライブ",
                         6: "RAM ディスク"}.get(dtype, "ドライブ")
            out.append(("%s (%s:)" % (label, chr(65 + i)), root))
    except Exception:
        pass
    return out


def user_folders():
    """デスクトップ/ドキュメント/ダウンロード/ピクチャ -> [(名前, パス)]"""
    home = os.path.expanduser("~")
    out = []
    for name, sub in (("デスクトップ", "Desktop"), ("ドキュメント", "Documents"),
                      ("ダウンロード", "Downloads"), ("ピクチャ", "Pictures")):
        p = os.path.join(home, sub)
        if os.path.isdir(p):
            out.append((name, p))
    return out


EDITOR_CANDS = (
    os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                 "sakura", "sakura.exe"),
    os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                 "sakura", "sakura.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "sakura",
                 "sakura.exe"),
    os.path.join(os.environ.get("APPDATA", ""), "sakura", "sakura.exe"),
    r"C:\sakura\sakura.exe",
)


def find_editor(conf_path=""):
    """テキストエディタ(既定はサクラエディタ)の実行ファイルを探す"""
    conf_path = (conf_path or "").strip().strip('"')
    if conf_path and os.path.isfile(conf_path):
        return conf_path
    for c in EDITOR_CANDS:
        if c and os.path.isfile(c):
            return c
    w = shutil.which("sakura.exe") or shutil.which("sakura")
    return w or ""


def open_terminal(kind, path):
    """path をカレントにして端末を開く。kind: cmd / powershell / wt
    UNCパスでも動くよう cmd は pushd、PowerShell は Set-Location を使う"""
    if os.name != "nt":
        return False
    path = path.rstrip("\\/") or path
    if len(path) == 2 and path[1] == ":":
        path += "\\"
    flag = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
    try:
        if kind == "cmd":
            subprocess.Popen('cmd.exe /k pushd "%s"' % path,
                             creationflags=flag)
        elif kind == "powershell":
            exe = "powershell.exe"
            cmd = "Set-Location -LiteralPath '%s'" % path.replace("'", "''")
            subprocess.Popen([exe, "-NoExit", "-NoLogo", "-Command", cmd],
                             creationflags=flag)
        elif kind == "wt":
            exe = shutil.which("wt.exe")
            if not exe:
                return False
            subprocess.Popen([exe, "-d", path])
        else:
            return False
        return True
    except OSError:
        return False


# ============================================================
#  検索語(フォルダ名/ファイル名検索共通)
#    スペース区切り=AND / 先頭 - で除外 / * ? を含む語はワイルドカード
# ============================================================
def parse_query(text):
    """検索欄の文字列 -> [(除外?, ワイルドカード?, 小文字語)]"""
    toks = []
    for w in text.split():
        neg = False
        if w.startswith("-") and len(w) > 1:
            neg = True
            w = w[1:]
        w = w.lower()
        if not w:
            continue
        # ワイルドカードは * ? だけ。[ ] は「[済]」のような名前の一部として扱う
        glob = "*" in w or "?" in w
        toks.append((neg, glob, w))
    return toks


@functools.lru_cache(maxsize=256)
def glob_regex(w):
    """ワイルドカード語 -> 名前全体に一致する正規表現(* と ? 以外は文字どおり)"""
    out = []
    for c in w:
        if c == "*":
            out.append(".*")
        elif c == "?":
            out.append(".")
        else:
            out.append(re.escape(c))
    return re.compile("".join(out) + r"\Z", re.S)


def match_word(hay, glob, w):
    return glob_regex(w).match(hay) is not None if glob else w in hay


def match_query(name_low, toks):
    """小文字化した名前が全条件を満たすか"""
    for neg, glob, w in toks:
        if match_word(name_low, glob, w) == neg:
            return False
    return True


def enable_file_drop(hwnd, on_drop):
    """WM_DROPFILESを購読し on_drop(paths) を呼ぶ。戻り値は要参照保持"""
    import ctypes
    WM_DROPFILES = 0x0233
    GWL_WNDPROC = -4
    u32 = ctypes.windll.user32
    s32 = ctypes.windll.shell32
    s32.DragAcceptFiles.argtypes = [ctypes.c_void_p, ctypes.c_int]
    s32.DragQueryFileW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                   ctypes.c_wchar_p, ctypes.c_uint]
    s32.DragFinish.argtypes = [ctypes.c_void_p]
    s32.DragAcceptFiles(hwnd, True)
    LRESULT = ctypes.c_ssize_t
    WNDPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_void_p, ctypes.c_uint,
                                 ctypes.c_void_p, ctypes.c_void_p)
    u32.CallWindowProcW.restype = LRESULT
    u32.CallWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                    ctypes.c_uint, ctypes.c_void_p,
                                    ctypes.c_void_p]
    SetWL = getattr(u32, "SetWindowLongPtrW", u32.SetWindowLongW)
    SetWL.restype = ctypes.c_void_p
    SetWL.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
    state = {}

    def proc(h, msg, wp, lp):
        if msg == WM_DROPFILES:
            try:
                n = s32.DragQueryFileW(wp, 0xFFFFFFFF, None, 0)
                paths = []
                for i in range(n):
                    ln = s32.DragQueryFileW(wp, i, None, 0)
                    buf = ctypes.create_unicode_buffer(ln + 1)
                    s32.DragQueryFileW(wp, i, buf, ln + 1)
                    paths.append(buf.value)
                s32.DragFinish(wp)
                on_drop(paths)
            except Exception:
                pass
            return 0
        return u32.CallWindowProcW(state["old"], h, msg, wp, lp)

    cb = WNDPROC(proc)
    state["old"] = SetWL(hwnd, GWL_WNDPROC,
                         ctypes.cast(cb, ctypes.c_void_p))
    return cb, state


def unique_dest(path):
    """コピー先が既に存在する場合 name (2).ext 形式で回避"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while True:
        cand = "%s (%d)%s" % (base, i, ext)
        if not os.path.exists(cand):
            return cand
        i += 1


# ============================================================
#  共通ユーティリティ
# ============================================================
def jp_font_families(root):
    jp = []
    jp_char = re.compile(r"[ぁ-んァ-ヺー々一-鿿ｦ-ﾟ]")
    for f in sorted(set(tkfont.families(root))):
        if f.startswith("@"):
            continue
        if jp_char.search(f) or any(k in f.lower() for k in JP_KEYS):
            jp.append(f)
    if not jp:
        jp = sorted(set(tkfont.families(root)))
    return jp


def pick_default_font(families):
    for cand in DEFAULT_FONTS:
        if cand in families:
            return cand
    return families[0] if families else "TkDefaultFont"


def dark_title_bar(root, dark):
    try:
        import ctypes
        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        val = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attr, ctypes.byref(val), 4)
    except Exception:
        pass


class Tip:
    """ボタンにマウスを乗せた時にショートカットを表示するツールチップ"""

    def __init__(self, wgt, text):
        self.wgt = wgt
        self.text = text
        self.tw = None
        wgt.bind("<Enter>", self.show)
        wgt.bind("<Leave>", self.hide)
        wgt.bind("<ButtonPress>", self.hide)

    def show(self, event=None):
        if self.tw:
            return
        x = self.wgt.winfo_rootx() + 8
        y = self.wgt.winfo_rooty() + self.wgt.winfo_height() + 4
        self.tw = tk.Toplevel(self.wgt)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        tk.Label(self.tw, text=self.text, bg="#ffffe1", fg="#000000",
                 relief="solid", borderwidth=1, padx=6, pady=2).pack()

    def hide(self, event=None):
        if self.tw:
            self.tw.destroy()
            self.tw = None


def fmt_size(n):
    if n < 0:
        return ""
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return ("%d %s" % (n, u)) if u == "B" else ("%.1f %s" % (n, u))
        n = n / 1024.0
    return "%.1f TB" % n


def fmt_time(t):
    if not t:
        return ""
    return time.strftime("%Y/%m/%d %H:%M", time.localtime(t))


def _tree_cols(tree):
    return ("#0",) + tuple(tree["columns"])


def col_widths(tree):
    """一覧の列幅 -> "#0:280,kind:100,..." (設定ファイル保存用)"""
    return ",".join("%s:%d" % (c, tree.column(c, "width"))
                    for c in _tree_cols(tree))


def apply_col_widths(tree, text):
    """col_widths() で保存した列幅を戻す(壊れた値は無視)"""
    cols = _tree_cols(tree)
    for item in text.split(","):
        c, _s, w = item.partition(":")
        if c in cols and w.isdigit():
            tree.column(c, width=max(int(w), 20))


def autofit_column(tree, col, font):
    """列見出しの境界をダブルクリック: 列幅を表示中の内容に合わせる"""
    if col not in _tree_cols(tree):
        return
    f = tkfont.Font(font=font)
    w = f.measure(tree.heading(col, "text")) + 30
    # 名前列はアイコン(16)と字下げぶんを足す
    extra = 50 if col == "#0" else 16
    for iid in tree.get_children()[:MAX_ROWS]:
        txt = tree.item(iid, "text") if col == "#0" else tree.set(iid, col)
        w = max(w, f.measure(txt) + extra)
    tree.column(col, width=min(w, 1600))


def load_favs():
    favs = []
    try:
        with open(FAV_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line:
                    k, v = line.split("=", 1)
                    if k.strip() and v.strip():
                        favs.append((k.strip(), v.strip()))
    except OSError:
        pass
    return favs


def fav_group(name):
    """お気に入り名を (グループ, 表示名) に分解。区切りは / または ／"""
    for sep in ("/", "／"):
        if sep in name:
            g, label = name.split(sep, 1)
            g = g.strip()
            label = label.strip()
            if g and label:
                return g, label
    return "", name


def save_favs(favs):
    try:
        with open(FAV_FILE, "w", encoding="utf-8") as f:
            for k, v in favs:
                f.write(k + "=" + v + "\n")
    except OSError:
        pass


def load_roots():
    roots = []
    try:
        with open(ROOTS_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip().strip('"')
                if line:
                    roots.append(line)
    except OSError:
        pass
    return roots


class ExpTab:
    """エクスプローラのタブ1枚分の状態"""

    def __init__(self, name="ホーム", path=None, key=""):
        self.name = name
        self.scope_dir = path
        self.scope = []
        self.hist_back = []
        self.hist_fwd = []
        self.sort_key = "name"
        self.sort_desc = False
        self.word = ""
        self.depth = 1
        self.search = None
        self.key = key
        self.custom = False


def load_tabs():
    tabs = []
    try:
        with open(TABS_FILE, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.rstrip("\n").split("\t")
                if not parts or not parts[0]:
                    continue
                t = ExpTab(parts[0],
                           parts[1] if len(parts) > 1 and parts[1] else None,
                           parts[2] if len(parts) > 2 else "")
                t.custom = len(parts) > 3 and parts[3] == "1"
                tabs.append(t)
    except OSError:
        pass
    if not tabs:
        tabs = [ExpTab()]
    return tabs


class Pane:
    """エクスプローラの1ペイン分の状態と部品"""

    def __init__(self, idx):
        self.idx = idx
        self.frame = None
        self.label = None
        self.var_label = None
        self.tree = None
        self.fpaths = {}
        self.fkind = {}
        self.scope_dir = None
        self.scope = []
        self.gen = 0
        self.hist_back = []
        self.hist_fwd = []
        self.sort_key = "name"
        self.sort_desc = False
        self.word = ""
        self.depth = 1            # 今の一覧を何層まで読んだか(再読込で同じ深さを使う)
        self.search = None        # フォルダ名検索の結果表示中なら (検索語, パス全体?)
        self.dtotal = 0           # フォルダ名検索の総ヒット数
        self.sel_target = None
        self.f_stop = threading.Event()


def _pane_prop(name):
    def g(self):
        return getattr(self.panes[self.cur_pane], name)

    def s(self, v):
        setattr(self.panes[self.cur_pane], name, v)
    return property(g, s)


# ============================================================
#  フォルダ名インデックス
# ============================================================
def norm_root(r):
    """検索ルートの表記をそろえる。末尾の区切りは落とすが、ドライブ直下は
    C:\\ の形を保つ(「C:」だとドライブ直下ではなくそのドライブの
    カレントフォルダを指してしまい、インデックスが空になるため)"""
    r = r.strip().strip('"')
    s = r.rstrip("\\/")
    if len(s) == 2 and s[1] == ":":
        return s + "\\"
    return s or r


def _is_junction(e):
    """ジャンクション(マウントポイント)なら True。中へは潜らない
    (Application Data 等の循環を避ける。OneDrive等のクラウドフォルダは対象外)"""
    if os.name != "nt":
        return False
    try:
        tag = getattr(e.stat(follow_symlinks=False), "st_reparse_tag", 0)
    except OSError:
        return False
    return tag == 0xA0000003


def scan_dirs(roots, stop, progress, files=None):
    """roots 以下のフォルダを全部列挙して返す。files にリストを渡すと
    ファイルのパスもそこへ集める(Excel等の ~$ 一時ファイルは除く)"""
    result = []
    work = queue.Queue()
    lock = threading.Lock()
    state = {"pending": 0, "done": 0}
    uniq = []
    for r in roots:
        r = norm_root(r)
        low = r.lower().rstrip("\\/")
        if r and os.path.isdir(r) and low not in [u[1] for u in uniq]:
            uniq.append((r, low))
    for r, low in uniq:
        # 別のルートの配下にあるルートは二重登録になるので除く
        if any(low != o and (low.startswith(o + "\\")
                             or low.startswith(o + "/"))
               for _r, o in uniq):
            continue
        state["pending"] += 1
        work.put(r)
        result.append(r)

    def worker():
        while True:
            if stop.is_set():
                return
            try:
                d = work.get(timeout=0.3)
            except queue.Empty:
                with lock:
                    if state["pending"] == 0:
                        return
                continue
            found = []
            subs = []
            fl = []
            try:
                with os.scandir(d) as it:
                    for e in it:
                        try:
                            if e.is_dir(follow_symlinks=False):
                                if e.name.lower() not in SKIP_DIRS:
                                    found.append(e.path)
                                    if not _is_junction(e):
                                        subs.append(e.path)
                            elif files is not None \
                                    and not is_office_temp(e.name):
                                fl.append(e.path)
                        except OSError:
                            pass
            except OSError:
                pass
            with lock:
                result.extend(found)
                if fl:
                    files.extend(fl)
                state["pending"] += len(subs) - 1
                state["done"] += 1
                done = state["done"]
                total = len(result)
            for s in subs:
                work.put(s)
            if done % 200 == 0:
                progress(done, total)

    ths = [threading.Thread(target=worker, daemon=True)
           for _ in range(SCAN_WORKERS)]
    for t in ths:
        t.start()
    for t in ths:
        t.join()
    return result


def save_dirindex(paths, path=None):
    path = path or DIRIDX_FILE
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", errors="replace") as f:
        for p in paths:
            f.write(p + "\n")
    os.replace(tmp, path)


def read_conf():
    """設定ファイル(doc_search_conf.txt) -> dict"""
    conf = {}
    try:
        with open(CONF_FILE, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.split("=", 1)
                    conf[k.strip()] = v.strip()
    except OSError:
        pass
    return conf


def _basename(p):
    return os.path.basename(p.rstrip("\\/")) or p


def parent_of(path):
    """フォルダのパス -> 親フォルダ(ドライブ直下などは "")"""
    s = path.rstrip("\\/")
    parent = os.path.dirname(s)
    return "" if not parent or parent == s else parent


def _line(blob, starts, li):
    """改行連結テキストの li 行目 -> (行の文字列, 行末位置)"""
    end = starts[li + 1] - 1 if li + 1 < len(starts) else len(blob)
    return blob[starts[li]:end], end


def _is_pathword(w):
    return "\\" in w or "/" in w


class _IndexSnap:
    """インデックスの中身一式。検索スレッドと再スキャンが並行しても
    食い違わないよう、作り直す時は丸ごと差し替える(中身は変更しない)"""

    def __init__(self, paths):
        starts = []
        parts = []
        nstarts = []
        nparts = []
        pos = npos = 0
        for p in paths:
            low = p.lower().replace("\n", " ")
            starts.append(pos)
            parts.append(low)
            pos += len(low) + 1
            nl = _basename(low)
            nstarts.append(npos)
            nparts.append(nl)
            npos += len(nl) + 1
        self.paths = paths
        self.blob = "\n".join(parts)     # パス全体(小文字)を改行連結
        self.starts = starts
        self.nblob = "\n".join(nparts)   # フォルダ名のみ(小文字)を改行連結
        self.nstarts = nstarts


class DirIndex:
    """フォルダ(またはファイル)の全パス一覧。名前(末尾要素)用とパス全体用の
    2つの検索用テキストを持ち、既定は名前そのものに一致させる"""

    def __init__(self, path=None):
        self.path = path            # None なら DIRIDX_FILE
        self.snap = _IndexSnap([])

    @property
    def paths(self):
        return self.snap.paths

    def load(self):
        paths = []
        try:
            with open(self.path or DIRIDX_FILE, "r", encoding="utf-8",
                      errors="replace") as f:
                for ln in f:
                    p = ln.rstrip("\r\n")
                    if not p.strip():
                        continue
                    # 旧版の不具合で入った「C:」「C:xxx」(ドライブ直下では
                    # なくカレントフォルダ基準の相対パス)を補正/除外
                    if re.match(r"^[A-Za-z]:$", p):
                        p += "\\"
                    elif re.match(r"^[A-Za-z]:[^\\/]", p):
                        continue
                    paths.append(p)
        except OSError:
            pass
        self.set_paths(paths)

    def set_paths(self, paths):
        self.snap = _IndexSnap(list(paths))

    def search(self, query, limit, full_path=False, sort="rank",
               desc=False, stop=None, with_keys=False):
        """query: 検索欄の文字列。full_path=False ならフォルダ名のみに一致
        (\\ や / を含む語だけはパス全体に照合)。
        sort: "rank"=関連度順(完全一致 > 前方一致 > 部分一致、浅い階層優先)
              "name"=名前順 / "path"=パス順
        戻り値: (総ヒット数, 並べ替え後の上位limit件のパス)。
        with_keys=True なら上位は (並べ替えキー, パス) のリスト
        (複数のインデックスの結果を同じ順序で混ぜるのに使う)。
        stop(Event)がセットされたら中断して None を返す"""
        s = self.snap
        toks = parse_query(query)
        if not s.paths or not toks:
            return 0, []
        # 語ごとに照合先を決める: True=パス全体 / False=フォルダ名
        conds = [(neg, glob, w, full_path or _is_pathword(w))
                 for neg, glob, w in toks]
        use_name = any(not onp for _n, _g, _w, onp in conds)
        use_path = any(onp for _n, _g, _w, onp in conds)

        def ok(name, path):
            for neg, glob, w, onp in conds:
                if match_word(path if onp else name, glob, w) == neg:
                    return False
            return True

        keys = [(w, onp) for neg, glob, w, onp in conds
                if not neg and not glob]
        hits = []
        if keys:
            # 一番長い語で候補行を高速に拾い、残りの条件は行ごとに確認
            key, onp = max(keys, key=lambda k: len(k[0]))
            blob, starts = (s.blob, s.starts) if onp else (s.nblob, s.nstarts)
            pos = 0
            n = 0
            while True:
                i = blob.find(key, pos)
                if i < 0:
                    break
                li = bisect.bisect_right(starts, i) - 1
                line, end = _line(blob, starts, li)
                name = line if not onp else (
                    _line(s.nblob, s.nstarts, li)[0] if use_name else "")
                path = line if onp else (
                    _line(s.blob, s.starts, li)[0] if use_path else "")
                if ok(name, path):
                    hits.append(li)
                pos = end + 1
                n += 1
                if not n & 0xFFF and stop is not None and stop.is_set():
                    return None
        else:
            names = s.nblob.split("\n") if use_name else None
            lows = s.blob.split("\n") if use_path else None
            for li in range(len(s.paths)):
                if ok(names[li] if use_name else "",
                      lows[li] if use_path else ""):
                    hits.append(li)
                if not li & 0xFFF and stop is not None and stop.is_set():
                    return None
        total = len(hits)
        first = next((w for neg, glob, w, onp in conds
                      if not neg and not glob and not _is_pathword(w)), "")

        def name_of(li):
            return _line(s.nblob, s.nstarts, li)[0]

        def path_of(li):
            return _line(s.blob, s.starts, li)[0]

        if sort == "name":
            def keyf(li):
                return (name_of(li), path_of(li))
        elif sort == "path":
            keyf = path_of
        else:
            desc = False

            def keyf(li):
                name = name_of(li)
                p = s.paths[li]
                if not first:
                    r = 2
                elif name == first or name.rsplit(".", 1)[0] == first:
                    r = 0        # 完全一致(ファイルは拡張子を除いて一致も)
                elif name.startswith(first):
                    r = 1
                elif first in name:
                    r = 2
                else:
                    r = 3        # パス側だけで一致
                return (r, p.count("\\") + p.count("/"), len(name),
                        path_of(li))

        # 件数を絞る前に全ヒットから順位付けする(完全一致が漏れないように)
        if total <= limit:
            top = sorted(hits, key=keyf, reverse=desc)
        elif desc:
            top = heapq.nlargest(limit, hits, key=keyf)
        else:
            top = heapq.nsmallest(limit, hits, key=keyf)
        if with_keys:
            return total, [(keyf(li), s.paths[li]) for li in top]
        return total, [s.paths[li] for li in top]


def is_office_temp(name):
    """Excel/Word等の編集中にできる一時ファイルか
    (~$見積書.xlsx などの所有者ファイル、~WRL0001.tmp などの作業ファイル)"""
    return name.startswith("~$") or (
        name.startswith("~") and name.lower().endswith(".tmp"))


def list_under(folder, depth, stop, out_q, pidx, gen):
    count = [0]

    def walk(d, lv):
        if stop.is_set() or count[0] >= MAX_LIST:
            return
        try:
            with os.scandir(d) as it:
                entries = list(it)
        except OSError:
            return
        for e in entries:
            if stop.is_set() or count[0] >= MAX_LIST:
                return
            try:
                isd = e.is_dir(follow_symlinks=False)
            except OSError:
                continue
            size = -1
            mt = 0
            try:
                stt = e.stat(follow_symlinks=False)
                mt = stt.st_mtime
                if not isd:
                    size = stt.st_size
            except OSError:
                pass
            count[0] += 1
            out_q.put(("item", pidx, gen,
                       "フォルダ" if isd else "ファイル",
                       e.name, e.path, size, mt))
            if isd and lv < depth:
                walk(e.path, lv + 1)

    walk(folder, 1)
    out_q.put(("listdone", pidx, gen, count[0]))


def cli_scan():
    roots = load_roots()
    if not roots:
        print("doc_search_roots.txt に対象ルートを1行1つで記載してください")
        return
    stop = threading.Event()

    def prog(done, total):
        if sys.stdout:
            sys.stdout.write("\rスキャン中... %d フォルダ" % total)
            sys.stdout.flush()

    print("対象ルート: " + ", ".join(roots))
    files = [] if read_conf().get("fileidx", "1") == "1" else None
    paths = scan_dirs(roots, stop, prog, files)
    save_dirindex(paths)
    if files is not None:
        save_dirindex(files, FILEIDX_FILE)
        print("\n完了: %d フォルダ / %d ファイルをインデックス化しました"
              % (len(paths), len(files)))
    else:
        print("\n完了: %d フォルダをインデックス化しました" % len(paths))


# ============================================================
#  文書内検索(抽出エンジン)
# ============================================================
def _fold_table():
    """全角英数記号・全角スペース・半角カナを、文字数を変えずに標準形へ
    寄せる表。長さが変わらないので、寄せた文章で見つけた位置のまま
    元の文章から前後を切り出せる"""
    tbl = {c: c - 0xFEE0 for c in range(0xFF01, 0xFF5F)}   # ！〜～ -> !〜~
    tbl[0x3000] = 0x20                                     # 全角スペース
    for c in range(0xFF61, 0xFFA0):                        # 半角カナ
        n = unicodedata.normalize("NFKC", chr(c))
        if len(n) == 1:
            tbl[c] = {"\u3099": "゛", "\u309a": "゜"}.get(n, n)
    return tbl


FOLD_TABLE = _fold_table()


def fold_text(text):
    return text if text.isascii() else text.translate(FOLD_TABLE)


def make_pattern(word, fold=True):
    """検索語 -> 正規表現(大文字小文字は区別しない)。
    fold=True なら全角/半角の違いも無視する(ガ は半角の ｶﾞ にも一致)"""
    if not fold:
        return re.compile(re.escape(word), re.I)
    w = fold_text(word).replace("゛", "\u3099").replace("゜", "\u309a")
    out = []
    for ch in unicodedata.normalize("NFC", w):
        d = unicodedata.normalize("NFD", ch)
        if len(d) == 2 and d[1] in "\u3099\u309a":
            # 濁音・半濁音は「カ゛」(半角カナを寄せた形)にも一致させる
            alt = d[0] + ("゛" if d[1] == "\u3099" else "゜")
            out.append("(?:%s|%s)" % (re.escape(ch), re.escape(alt)))
        else:
            out.append(re.escape(ch))
    return re.compile("".join(out), re.I)


def snippets(text, pat, ctx, fold=True):
    """text 中の pat の出現ごとに前後 ctx 文字を切り出す(最大 MAX_HIT 件)"""
    res = []
    hay = fold_text(text) if fold else text
    for m in pat.finditer(hay):
        s = max(0, m.start() - ctx)
        e = min(len(text), m.end() + ctx)
        res.append(re.sub(r"\s+", " ", text[s:e]))
        if len(res) >= MAX_HIT:
            break
    return res


def _natkey(s):
    """sheet2 < sheet10 となる並べ替えキー"""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]


def join_runs(xml, tag, para_end):
    lines = []
    # タグ名の後ろは空白か > のみ(<w:t> が <w:tab/> 等に誤一致しないように)
    pat = re.compile("<" + tag + r"(?:\s[^>]*)?>([^<]*)</" + tag + ">")
    for para in xml.split(para_end):
        t = "".join(pat.findall(para))
        if t:
            lines.append(html.unescape(t))
    return "\n".join(lines)


DOCX_PARTS = (
    (r"word/document\.xml$", "本文"),
    (r"word/(header|footer)\d+\.xml$", "ヘッダ/フッタ"),
    (r"word/footnotes\.xml$", "脚注"),
    (r"word/endnotes\.xml$", "文末脚注"),
    (r"word/comments\.xml$", "コメント"),
)


def extract_docx(zf):
    parts = []
    names = sorted(zf.namelist(), key=_natkey)
    for pat, label in DOCX_PARTS:
        for name in names:
            if not re.match(pat, name):
                continue
            xml = zf.read(name).decode("utf-8", "ignore")
            text = join_runs(xml, "w:t", "</w:p>")
            if text:
                parts.append((label, text))
    return parts


def extract_pptx(zf):
    parts = []
    for name in sorted(zf.namelist(), key=_natkey):
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
        n = re.match(r"ppt/notesSlides/notesSlide(\d+)\.xml$", name)
        if not (m or n):
            continue
        xml = zf.read(name).decode("utf-8", "ignore")
        text = join_runs(xml, "a:t", "</a:p>")
        if text:
            label = "スライド" + m.group(1) if m else "ノート" + n.group(1)
            parts.append((label, text))
    return parts


XL_T = re.compile(r"<t(?:\s[^>]*)?>([^<]*)</t>")


def parse_sst(zf):
    sst = []
    try:
        xml = zf.read("xl/sharedStrings.xml").decode("utf-8", "ignore")
    except KeyError:
        return sst
    xml = re.sub(r"<rPh\b.*?</rPh>", "", xml, flags=re.S)
    for item in re.findall(r"<si>(.*?)</si>|<si/>", xml, flags=re.S):
        t = "".join(XL_T.findall(item))
        sst.append(html.unescape(t))
    return sst


def xlsx_sheet_names(zf):
    names = {}
    try:
        wb = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    except KeyError:
        return names
    rid2file = {}
    for tag in re.findall(r"<Relationship [^>]*>", rels):
        mi = re.search(r'Id="([^"]+)"', tag)
        mt = re.search(r'Target="([^"]+)"', tag)
        if mi and mt:
            rid2file[mi.group(1)] = mt.group(1).split("/")[-1]
    for tag in re.findall(r"<sheet [^>]*>", wb):
        mn = re.search(r'name="([^"]*)"', tag)
        mr = re.search(r'r:id="([^"]+)"', tag)
        if mn and mr and mr.group(1) in rid2file:
            names[rid2file[mr.group(1)]] = html.unescape(mn.group(1))
    return names


def xlsx_drawing_map(zf, names):
    """図形(drawingN.xml)・コメント(commentsN.xml) -> 置かれているシート名"""
    dmap = {}
    for name in zf.namelist():
        m = re.match(r"xl/worksheets/_rels/(sheet\d+\.xml)\.rels$", name)
        if not m:
            continue
        rels = zf.read(name).decode("utf-8", "ignore")
        for mt in re.findall(
                r'Target="[^"]*?((?:drawing|comments)\d+\.xml)"', rels):
            dmap[mt] = names.get(m.group(1), m.group(1))
    return dmap


# セル: <c ...>...</c> と空セル <c .../> の両方に対応
# (空セルを取りこぼすと次のセルとくっついて値を読み違える)
XL_CELL = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
XL_TYPE = re.compile(r'\bt="(\w+)"')
XL_V = re.compile(r"<v>([^<]*)</v>")


def extract_sheet(xml, sst):
    texts = []
    for attrs, body in XL_CELL.findall(xml):
        if not body:
            continue
        m = XL_TYPE.search(attrs)
        t = m.group(1) if m else "n"
        if t == "inlineStr":
            s = "".join(XL_T.findall(body))
            if s:
                texts.append(html.unescape(s))
            continue
        v = XL_V.search(body)
        if not v or not v.group(1):
            continue
        if t == "s":
            try:
                i = int(v.group(1))
            except ValueError:
                continue
            if i < len(sst) and sst[i]:
                texts.append(sst[i])
        elif t in ("str", "n", "d"):
            # 数式の文字列結果 / 数値 / 日付(ISO形式)。数値もExcelの検索同様に対象
            texts.append(html.unescape(v.group(1)))
    return "\n".join(texts)


def extract_xlsx(zf):
    parts = []
    sst = parse_sst(zf)
    names = xlsx_sheet_names(zf)
    dmap = xlsx_drawing_map(zf, names)
    order = {f: i for i, f in enumerate(names)}    # ブック上のシートの並び

    def key(name):
        base = name.rsplit("/", 1)[-1]
        return (order.get(base, len(order)), _natkey(name))

    for name in sorted(zf.namelist(), key=key):
        m = re.match(r"xl/worksheets/(sheet\d+\.xml)$", name)
        d = re.match(r"xl/drawings/(drawing\d+\.xml)$", name)
        c = re.match(r"xl/(comments\d+\.xml)$", name)
        if m:
            xml = zf.read(name).decode("utf-8", "ignore")
            text = extract_sheet(xml, sst)
            if text:
                parts.append((names.get(m.group(1), m.group(1)), text))
        elif d or c:
            xml = zf.read(name).decode("utf-8", "ignore")
            if d:
                text = join_runs(xml, "a:t", "</a:p>")
                kind = "図形"
            else:
                text = join_runs(xml, "t", "</comment>")
                kind = "コメント"
            if text:
                sheet = dmap.get((d or c).group(1), "")
                label = kind + "(" + sheet + ")" if sheet else kind
                parts.append((label, text))
    return parts


def read_text(path):
    with open(path, "rb") as f:
        data = f.read()
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16", "replace")      # BOM付きUTF-16(Unicodeテキスト)
    for enc in ("utf-8-sig", "cp932"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("cp932", "replace")


def search_worker(folder, word, ctx, excludes, exts, q, stop=None,
                  fold=True):
    n_files = n_hits = n_err = 0
    pat = make_pattern(word, fold)
    for root, dirs, files in os.walk(folder):
        if stop is not None and stop.is_set():
            break
        dirs[:] = [d for d in dirs
                   if not d.startswith((".", "$"))
                   and not any(x in os.path.join(root, d).lower()
                               for x in excludes)]
        for fn in files:
            if stop is not None and stop.is_set():
                break
            if fn.startswith("~$"):
                continue
            path = os.path.join(root, fn)
            if any(x in path.lower() for x in excludes):
                continue
            ext = os.path.splitext(fn)[1].lower()
            try:
                if ext in (".docx", ".docm"):
                    with zipfile.ZipFile(path) as zf:
                        parts = extract_docx(zf)
                elif ext in (".xlsx", ".xlsm"):
                    with zipfile.ZipFile(path) as zf:
                        parts = extract_xlsx(zf)
                elif ext in (".pptx", ".pptm"):
                    with zipfile.ZipFile(path) as zf:
                        parts = extract_pptx(zf)
                elif ext in exts:
                    parts = [("", read_text(path))]
                else:
                    continue
                n_files += 1
                folder_path = os.path.dirname(path)
                for label, text in parts:
                    for frag in snippets(text, pat, ctx, fold):
                        n_hits += 1
                        q.put(("hit", path, fn, folder_path, label, frag))
            except Exception as e:
                n_err += 1
                q.put(("err", path, fn, os.path.dirname(path),
                       "エラー", str(e)))
    q.put(("done", n_files, n_hits, n_err))


# ============================================================
#  GUI
# ============================================================
class App:
    # アクティブなペインへ委譲するプロパティ群
    scope_dir = _pane_prop("scope_dir")
    scope = _pane_prop("scope")
    hist_back = _pane_prop("hist_back")
    hist_fwd = _pane_prop("hist_fwd")
    sort_key = _pane_prop("sort_key")
    sort_desc = _pane_prop("sort_desc")
    sel_target = _pane_prop("sel_target")
    fpaths = _pane_prop("fpaths")
    fkind = _pane_prop("fkind")
    f_stop = _pane_prop("f_stop")

    @property
    def ftree(self):
        return self.panes[self.cur_pane].tree

    def __init__(self, root):
        self.root = root
        root.title("文書横断検索")
        root.geometry("1180x680")
        self.app_icons = [rows_to_photo(app_icon_rows(s))
                          for s in (48, 32, 16)]
        root.iconphoto(True, *self.app_icons)
        self.q = queue.Queue()
        self.fq = queue.Queue()
        self.paths = {}
        self.running = False
        self.scanning = False
        self.hit_count = 0
        self.row_i = 0
        self.last_open = 0.0
        self.menu_iid = ""
        self.fmenu_iid = ""
        self.tmenu_idx = 0
        self.favs = load_favs()
        self.index = DirIndex()
        self.idx_state = "none"
        self.findex = DirIndex(FILEIDX_FILE)    # ファイル名のインデックス
        self.fidx_state = "none"
        self.etabs = load_tabs()
        self.cur_et = 0
        self.panes = [Pane(0), Pane(1)]
        self.cur_pane = 0
        self.pane_by_tree = {}
        self.gen_seq = 0
        self.focus_list = False
        self.sel_memory = {}
        self._side_job = None
        self._drag_iid = None
        self._drag_started = False
        self._press_y = 0
        self.last_group = ""
        self._dropkeep = []
        self._pending_fsearch = False
        self._ta_buf = ""
        self._ta_time = 0.0
        self.scan_stop = threading.Event()
        self.doc_stop = threading.Event()
        # フォルダ検索タブ(Everything風)の状態
        self.find_paths = {}
        self.find_isfile = {}
        self.find_sort = ("rank", False)
        self._find_gen = 0
        self._find_stop = threading.Event()
        self._find_job = None
        self._find_shown = None       # 表示中の結果の検索条件
        self._find_running = None     # 実行中の検索の検索条件
        self._find_focus_after = False
        self.dmenu_iid = ""
        self.style = ttk.Style(root)
        self.families = jp_font_families(root)
        self.icons = build_icons()
        conf = self.load_conf()
        self.two_pane = conf.get("panes", "1") == "2"

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True)
        self.tab1 = ttk.Frame(self.nb)
        self.tabf = ttk.Frame(self.nb)
        self.tabd = ttk.Frame(self.nb)
        tab2 = ttk.Frame(self.nb)
        self.nb.add(self.tab1, text=" 文書内検索 ")
        self.nb.add(self.tabf, text=" エクスプローラ ")
        self.nb.add(self.tabd, text=" フォルダ検索 ")
        self.nb.add(tab2, text=" 設定 ")
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_change)
        tab1 = self.tab1

        # ================= 文書内検索タブ =================
        top = ttk.Frame(tab1, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="フォルダ:").grid(row=0, column=0, sticky="w")
        self.var_dir = tk.StringVar()
        self.ent_dir = ttk.Entry(top, textvariable=self.var_dir)
        self.ent_dir.grid(row=0, column=1, sticky="we", padx=6)
        self.ent_dir.bind("<Return>",
                          lambda e: self.focus_widget(self.ent_word))
        ttk.Button(top, text="参照(Ctrl+D)", command=self.browse).grid(
            row=0, column=2)
        self.btn_fav = ttk.Button(top, text="お気に入り(F7)",
                                  command=self.show_fav_menu)
        self.btn_fav.grid(row=0, column=3, padx=6)
        ttk.Button(top, text="追加(F8)", command=self.add_favorite).grid(
            row=0, column=4)
        ttk.Label(top, text="検索文字列:").grid(row=1, column=0, sticky="w",
                                          pady=6)
        wrap = ttk.Frame(top)
        wrap.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        self.var_word = tk.StringVar()
        self.ent_word = ttk.Entry(wrap, textvariable=self.var_word, width=30)
        self.ent_word.pack(side="left")
        self.ent_word.bind("<Return>", lambda e: self.start())
        ttk.Label(wrap, text="  前後").pack(side="left")
        self.var_ctx = tk.StringVar(value=str(CTX))
        self.spin = tk.Spinbox(wrap, from_=1, to=500,
                               textvariable=self.var_ctx, width=5)
        self.spin.pack(side="left", padx=2)
        ttk.Label(wrap, text="文字").pack(side="left")
        self.var_fold = tk.BooleanVar(value=conf.get("fold", "1") == "1")
        chk_fold = ttk.Checkbutton(wrap, text="全角/半角を区別しない",
                                   variable=self.var_fold,
                                   command=self.save_conf)
        chk_fold.pack(side="left", padx=(12, 0))
        Tip(chk_fold, "ON: ＡＢＣ と ABC、ｱｲｳ と アイウ を同じ文字とみなす\n"
                      "(大文字小文字はいつも区別しません)")
        self.btn = ttk.Button(top, text="検索(F5)", style="Accent.TButton",
                              command=self.toggle_doc_search)
        self.btn.grid(row=1, column=2, pady=6)
        ttk.Button(top, text="CSV出力(F9)", command=self.export_csv).grid(
            row=1, column=3, padx=6)
        ttk.Label(top, text="除外文字列:").grid(row=2, column=0, sticky="w")
        wrap2 = ttk.Frame(top)
        wrap2.grid(row=2, column=1, sticky="w", padx=6)
        self.var_ex = tk.StringVar()
        self.ent_ex = ttk.Entry(wrap2, textvariable=self.var_ex, width=30)
        self.ent_ex.pack(side="left")
        self.ent_ex.bind("<Return>", lambda e: self.start())
        ttk.Label(wrap2, style="Sub.TLabel",
                  text="  カンマ区切りで複数可・パスに含まれると除外"
                  ).pack(side="left")
        top.columnconfigure(1, weight=1)

        gold1 = tk.Frame(tab1, height=1)
        gold1.pack(fill="x", padx=10)

        mid = ttk.Frame(tab1)
        mid.pack(fill="both", expand=True, padx=10, pady=(6, 8))
        self.tree = ttk.Treeview(mid, columns=("dir", "where", "ctx"),
                                 show="tree headings")
        self.tree.heading("#0", text="ファイル名")
        self.tree.heading("dir", text="フォルダ")
        self.tree.heading("where", text="場所")
        self.tree.heading("ctx", text="前後の文脈")
        # 列幅: 伸縮(stretch)する列はマウスを離した時にTkが幅を配り直して
        # しまうため、手で変える列は stretch=False。余白は最後の文脈列で吸収
        self.tree.column("#0", width=200, minwidth=60, stretch=False)
        self.tree.column("dir", width=250, minwidth=60, stretch=False)
        self.tree.column("where", width=100, minwidth=40, anchor="center",
                         stretch=False)
        self.tree.column("ctx", width=420, minwidth=150, stretch=True)
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="we")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.on_dblclick)
        self.tree.bind("<Button-3>", self.on_rclick)
        self.tree.bind("<Return>", self.on_tree_folder)
        self.tree.bind("<Control-Return>", self.on_tree_file)
        self.tree.bind("<App>", self.on_menu_key)
        self.tree.bind("<Shift-F10>", self.on_menu_key)
        for w in (self.ent_dir, self.ent_word, self.ent_ex, self.tree):
            w.bind("<Escape>", self.on_doc_escape)

        self.menu = tk.Menu(root, tearoff=0)
        self.menu.add_command(label="フォルダを開く (Enter)",
                              command=lambda: self.open_folder_of(
                                  self.menu_iid))
        self.menu.add_command(label="ファイルを開く (ダブルクリック / Ctrl+Enter)",
                              command=lambda: self.open_file_of(
                                  self.menu_iid))
        self.menu.add_separator()
        self.menu.add_command(label="フルパスをコピー",
                              command=lambda: self.copy_path("full"))
        self.menu.add_command(label="フォルダパスをコピー",
                              command=lambda: self.copy_path("dir"))
        self.menu.add_command(label="ファイル名をコピー",
                              command=lambda: self.copy_path("name"))
        self.menu.add_separator()
        self.menu.add_command(label="サクラエディタで開く (Ctrl+Shift+E)",
                              command=lambda: self.open_in_editor(
                                  [self.paths.get(self.menu_iid, "")]))
        self.menu.add_command(label="このアプリのエクスプローラで表示",
                              command=lambda: self.reveal_in_explorer(
                                  self.menu_iid))
        self.menu.add_command(label="このフォルダでコマンドプロンプト",
                              command=lambda: self.open_terminal_here(
                                  "cmd", self.doc_dir_of(self.menu_iid)))
        self.menu.add_command(label="このフォルダでPowerShell",
                              command=lambda: self.open_terminal_here(
                                  "powershell",
                                  self.doc_dir_of(self.menu_iid)))

        # ================= エクスプローラタブ =================
        self.tabbar = ttk.Frame(self.tabf, padding=(10, 6, 10, 0))
        self.tabbar.pack(fill="x")

        ftop = ttk.Frame(self.tabf, padding=(10, 6, 10, 6))
        ftop.pack(fill="x")
        b_back = ttk.Button(ftop, text="←", width=3, style="Tool.TButton",
                            command=self.go_back)
        b_back.pack(side="left")
        Tip(b_back, "戻る (Alt+← / BackSpace)")
        b_fwd = ttk.Button(ftop, text="→", width=3, style="Tool.TButton",
                           command=self.go_fwd)
        b_fwd.pack(side="left", padx=(2, 0))
        Tip(b_fwd, "進む (Alt+→)")
        b_up = ttk.Button(ftop, text="↑", width=3, style="Tool.TButton",
                          command=self.go_up)
        b_up.pack(side="left", padx=(2, 0))
        Tip(b_up, "一つ上へ (Alt+↑)")
        b_home = ttk.Button(ftop, text="⌂", width=3, style="Tool.TButton",
                            command=self.go_home)
        b_home.pack(side="left", padx=(2, 8))
        Tip(b_home, "ホーム (Alt+Home)")
        self.var_addr = tk.StringVar()
        self.ent_addr = ttk.Entry(ftop, textvariable=self.var_addr)
        self.ent_addr.pack(side="left", fill="x", expand=True)
        self.ent_addr.bind("<Return>", self.on_addr_enter)
        self.ent_addr.bind("<Escape>", self.on_entry_escape)
        Tip(self.ent_addr, "パスを入力してEnter。cmd / powershell / wt と"
                           "入力するとこの場所で端末を開く")
        ttk.Label(ftop, text=" フォルダ名").pack(side="left")
        self.var_dword = tk.StringVar()
        self.ent_dword = ttk.Entry(ftop, textvariable=self.var_dword,
                                   width=13)
        self.ent_dword.pack(side="left", padx=(4, 0))
        self.ent_dword.bind("<Return>", self.fsearch_folder)
        self.ent_dword.bind("<Escape>", self.on_entry_escape)
        Tip(self.ent_dword, "全フォルダから名前で探す (Ctrl+F)\n"
                            "スペース=AND  -語=除外  *?=ワイルドカード")
        self.var_dpath = tk.BooleanVar(value=conf.get("dpath", "0") == "1")
        self.chk_dpath = ttk.Checkbutton(ftop, text="パス",
                                         variable=self.var_dpath,
                                         command=self.on_dpath_toggle)
        self.chk_dpath.pack(side="left", padx=(2, 0))
        Tip(self.chk_dpath, "ON: パス全体に含まれれば一致\n"
                            "OFF: フォルダ名そのものに含まれる場合だけ一致")
        ttk.Label(ftop, text=" ファイル名").pack(side="left")
        self.var_fword = tk.StringVar()
        self.ent_fword = ttk.Entry(ftop, textvariable=self.var_fword,
                                   width=13)
        self.ent_fword.pack(side="left", padx=(4, 0))
        self.ent_fword.bind("<Return>", self.fsearch_file)
        self.ent_fword.bind("<Escape>", self.on_entry_escape)
        Tip(self.ent_fword, "今の場所から深さN層の名前を探す (Ctrl+E)\n"
                            "スペース=AND  -語=除外  *.xlsx などワイルドカード可\n"
                            "空欄でEnterなら通常の一覧に戻る")
        ttk.Label(ftop, text=" 深さ").pack(side="left")
        self.var_depth = tk.StringVar(value=str(LIST_DEPTH))
        self.dspin = tk.Spinbox(ftop, from_=1, to=10, width=3,
                                textvariable=self.var_depth)
        self.dspin.pack(side="left", padx=2)
        # takefocus=False: クリックしても入力欄のフォーカスを奪わない。
        # 奪うと「どちらの欄で検索したか」が分からず、フォルダ名を入れて
        # 押してもファイル名検索(=深さN層の全件表示)が走ってしまう
        b_fs = ttk.Button(ftop, text="検索(F5)", style="Accent.TButton",
                          takefocus=False, command=self.fsearch)
        b_fs.pack(side="left", padx=(6, 0))
        Tip(b_fs, "カーソルのある欄(フォルダ名/ファイル名)で検索")
        self.var_showtmp = tk.BooleanVar(
            value=conf.get("showtmp", "0") == "1")
        chk_tmp = ttk.Checkbutton(ftop, text="~$表示", takefocus=False,
                                  variable=self.var_showtmp,
                                  command=self.on_showtmp_toggle)
        chk_tmp.pack(side="left", padx=(8, 0))
        Tip(chk_tmp, "Excel/Word等の編集中にできる一時ファイル\n"
                     "(~$で始まるファイル等)を表示する (Ctrl+H)")

        gold3 = tk.Frame(self.tabf, height=1)
        gold3.pack(fill="x", padx=10)

        paned = ttk.PanedWindow(self.tabf, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=(6, 8))
        sidef = ttk.Frame(paned)
        mainf = ttk.Frame(paned)
        paned.add(sidef, weight=0)
        paned.add(mainf, weight=1)

        self.var_sfilter = tk.StringVar()
        self.ent_side = ttk.Entry(sidef, textvariable=self.var_sfilter)
        self.ent_side.pack(fill="x", pady=(0, 2))
        Tip(self.ent_side, "お気に入り絞り込み (Ctrl+B)")
        self.ent_side.bind("<KeyRelease>", self.on_side_filter)
        self.ent_side.bind("<Escape>", self.clear_side_filter)
        self.ent_side.bind("<Return>",
                           lambda e: (self.focus_side(), "break")[1])
        self.ent_side.bind("<Down>",
                           lambda e: (self.focus_side(), "break")[1])
        self.side = ttk.Treeview(sidef, show="tree", columns=("p", "i"),
                                 displaycolumns=(),
                                 style="Side.Treeview", selectmode="browse")
        self.side.pack(fill="both", expand=True)
        self.side.bind("<ButtonPress-1>", self.on_side_press)
        self.side.bind("<B1-Motion>", self.on_side_motion)
        self.side.bind("<ButtonRelease-1>", self.on_side_release)
        self.side.bind("<Return>", self.on_side_enter)
        self.side.bind("<Control-Up>",
                       lambda e: self.on_side_move_key(-1))
        self.side.bind("<Control-Down>",
                       lambda e: self.on_side_move_key(1))

        self.head_base = {"name": "名前", "kind": "種類", "size": "サイズ",
                          "mtime": "更新日時", "path": "パス"}
        self.colid = {"name": "#0", "kind": "kind", "size": "size",
                      "mtime": "mtime", "path": "path"}
        self.mains = ttk.PanedWindow(mainf, orient="horizontal")
        self.mains.pack(fill="both", expand=True)
        for p in self.panes:
            self.build_pane(self.mains, p)
            self.pane_by_tree[p.tree] = p
        self.apply_pane_layout()

        self.fmenu = tk.Menu(root, tearoff=0)
        self.fmenu.add_command(label="開く (Enter)",
                               command=lambda: self.fm_do("open"))
        self.fmenu.add_command(label="Explorerで場所を開く (Ctrl+Enter)",
                               command=lambda: self.fm_do("place"))
        self.fmenu.add_command(label="新しいタブで開く",
                               command=lambda: self.fm_do("newtab"))
        self.fmenu.add_separator()
        self.fmenu.add_command(label="サクラエディタで開く (Ctrl+Shift+E)",
                               command=lambda: self.fm_do("editor"))
        self.fmenu.add_command(label="プログラムから開く...",
                               command=lambda: self.fm_do("openas"))
        self.fmenu_term = tk.Menu(self.fmenu, tearoff=0)
        self.fmenu_term.add_command(
            label="コマンドプロンプト (Ctrl+Shift+C)",
            command=lambda: self.fm_do("term_cmd"))
        self.fmenu_term.add_command(
            label="PowerShell (Ctrl+Shift+S)",
            command=lambda: self.fm_do("term_powershell"))
        self.fmenu_term.add_command(
            label="Windows Terminal",
            command=lambda: self.fm_do("term_wt"))
        self.fmenu.add_cascade(label="この場所で端末を開く",
                               menu=self.fmenu_term)
        self.fmenu.add_separator()
        self.fmenu.add_command(label="切り取り (Ctrl+X)",
                               command=self.cut_sel)
        self.fmenu.add_command(label="コピー (Ctrl+C)",
                               command=self.copy_sel)
        self.fmenu.add_command(label="貼り付け (Ctrl+V)",
                               command=self.paste_here)
        self.fmenu.add_command(label="ごみ箱へ削除 (Delete)",
                               command=self.delete_sel)
        self.fmenu.add_command(label="完全に削除 (Shift+Delete)",
                               command=lambda: self.delete_sel(True))
        self.fmenu.add_separator()
        self.fmenu.add_command(label="名前の変更 (F2)",
                               command=self.rename_sel)
        self.fmenu.add_command(label="新しいフォルダー (Ctrl+Shift+N)",
                               command=self.new_folder)
        self.fmenu.add_command(label="新しいテキスト ファイル",
                               command=self.new_text_file)
        self.fmenu.add_command(label="プロパティ (Alt+Enter)",
                               command=lambda: self.fm_do("props"))
        self.fmenu.add_separator()
        self.fmenu.add_command(label="お気に入りに追加 (F8)",
                               command=lambda: self.fm_do("fav"))
        self.fmenu.add_command(label="この場所で文書内検索 (Ctrl+G)",
                               command=lambda: self.fm_do("todoc"))
        self.fmenu.add_command(label="このフォルダ以下を再スキャン",
                               command=lambda: self.fm_do("rescan"))
        self.fmenu.add_separator()
        self.fmenu.add_command(label="フルパスをコピー",
                               command=lambda: self.fm_do("cp_full"))
        self.fmenu.add_command(label="フォルダパスをコピー",
                               command=lambda: self.fm_do("cp_dir"))
        self.fmenu.add_command(label="名前をコピー",
                               command=lambda: self.fm_do("cp_name"))

        self.tmenu = tk.Menu(root, tearoff=0)
        self.tmenu.add_command(label="名前を変更",
                               command=lambda: self.tab_act("rename"))
        self.tmenu.add_command(label="ショートカット設定 (Ctrl+数字)",
                               command=lambda: self.tab_act("key"))
        self.tmenu.add_command(label="タブを複製",
                               command=lambda: self.tab_act("dup"))
        self.tmenu.add_separator()
        self.tmenu.add_command(label="タブを閉じる (Ctrl+W)",
                               command=lambda: self.tab_act("close"))

        # ================= フォルダ検索タブ(Everything風) =================
        gold4 = self.build_find_tab(conf)

        # ================= 設定タブ =================
        trow = ttk.Frame(tab2)
        trow.pack(anchor="w", padx=14, pady=(12, 2))
        ttk.Label(trow, text="テーマ: ").pack(side="left")
        theme = conf.get("theme", "")
        if theme not in THEMES:
            theme = WIN_STD
        self._last_theme = theme
        self.var_theme = tk.StringVar(value=theme)
        self.cb_theme = ttk.Combobox(trow, textvariable=self.var_theme,
                                     width=18, state="readonly",
                                     values=list(THEMES))
        self.cb_theme.pack(side="left")
        self.cb_theme.bind("<<ComboboxSelected>>", self.on_conf_change)

        frow = ttk.Frame(tab2)
        frow.pack(anchor="w", padx=14, pady=2)
        ttk.Label(frow, text="フォント: ").pack(side="left")
        font = conf.get("font", "")
        if font not in self.families:
            font = pick_default_font(self.families)
        self.var_font = tk.StringVar(value=font)
        cb_font = ttk.Combobox(frow, textvariable=self.var_font, width=26,
                               state="readonly", values=self.families)
        cb_font.pack(side="left")
        cb_font.bind("<<ComboboxSelected>>", self.on_conf_change)
        ttk.Label(frow, text="  サイズ").pack(side="left")
        self.var_fsize = tk.StringVar(
            value=conf.get("size", "9" if theme == WIN_STD else "10"))
        self.fspin = tk.Spinbox(frow, from_=8, to=20, width=4,
                                textvariable=self.var_fsize,
                                command=self.on_conf_change)
        self.fspin.pack(side="left", padx=2)
        self.fspin.bind("<Return>", self.on_conf_change)

        erow = ttk.Frame(tab2)
        erow.pack(anchor="w", padx=14, pady=(8, 2), fill="x")
        ttk.Label(erow, text="テキストエディタ: ").pack(side="left")
        self.var_editor = tk.StringVar(value=conf.get("editor", ""))
        self.ent_editor = ttk.Entry(erow, textvariable=self.var_editor,
                                    width=52)
        self.ent_editor.pack(side="left")
        self.ent_editor.bind("<Return>", self.on_editor_change)
        self.ent_editor.bind("<FocusOut>", self.on_editor_change)
        ttk.Button(erow, text="参照", command=self.browse_editor).pack(
            side="left", padx=6)
        self.var_editor_note = tk.StringVar()
        ttk.Label(tab2, textvariable=self.var_editor_note,
                  style="Sub.TLabel", padding=(8, 0)).pack(anchor="w")
        self.update_editor_note()

        ttk.Label(tab2, padding=(8, 10, 8, 0),
                  text="テキストとして検索する拡張子（1行に1つ。ドット付き）"
                  ).pack(anchor="w")
        self.ext_text = tk.Text(tab2, width=24, height=7)
        self.ext_text.pack(anchor="w", padx=14)
        ttk.Button(tab2, text="拡張子を保存 (Ctrl+Shift+K)",
                   command=self.save_ext).pack(
            anchor="w", padx=14, pady=4)

        ttk.Label(tab2, padding=(8, 10, 8, 0),
                  text="お気に入り（1行に1つ。書式: 名前=パス または "
                       "グループ名/名前=パス ※ファイルも可）"
                  ).pack(anchor="w")
        self.fav_text = tk.Text(tab2, width=70, height=4)
        self.fav_text.pack(anchor="w", padx=14)
        ttk.Button(tab2, text="お気に入りを保存 (Ctrl+Shift+O)",
                   command=self.save_fav_edit).pack(
            anchor="w", padx=14, pady=4)

        ttk.Label(tab2, padding=(8, 10, 8, 0),
                  text="フォルダ名検索/フォルダ検索タブの検索ルート（1行に1つ。例: "
                       r"C:\ や \\server\share や D:\projects）"
                  ).pack(anchor="w")
        self.roots_text = tk.Text(tab2, width=70, height=3)
        self.roots_text.pack(anchor="w", padx=14)
        srow = ttk.Frame(tab2)
        srow.pack(anchor="w", padx=14, pady=4)
        ttk.Button(srow, text="インデックス作成/更新 (Ctrl+Shift+I)",
                   command=self.full_scan).pack(side="left")
        ttk.Button(srow, text="スキャン停止 (Ctrl+Shift+P)",
                   command=self.stop_scan).pack(
            side="left", padx=6)
        self.var_scandate = tk.StringVar(value="")
        ttk.Label(srow, textvariable=self.var_scandate,
                  style="Sub.TLabel").pack(side="left", padx=6)
        self.var_fileidx = tk.BooleanVar(
            value=conf.get("fileidx", "1") == "1")
        ttk.Checkbutton(tab2, text="ファイル名もインデックスに含める"
                                   "（フォルダ検索タブでファイルも探せる）",
                        variable=self.var_fileidx,
                        command=self.save_conf).pack(anchor="w", padx=14)
        ttk.Label(tab2, style="Sub.TLabel", padding=(8, 2),
                  text="夜間の自動更新: タスクスケジューラに "
                       "python doc_search.py --scan を登録"
                  ).pack(anchor="w")

        content = ""
        try:
            with open(EXT_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
        except OSError:
            pass
        if not content:
            content = "\n".join(TEXT_EXT)
        self.ext_text.insert("1.0", content)
        self.refresh_fav_text()
        self.roots_text.insert("1.0", "\n".join(load_roots()))

        gold2 = tk.Frame(root, height=1)
        gold2.pack(fill="x")
        self.gold_lines = [gold1, gold2, gold3, gold4]
        self.spins = [self.spin, self.fspin, self.dspin]
        self.texts = [self.ext_text, self.fav_text, self.roots_text]
        self.font_widgets = [self.ent_dir, self.ent_word, self.ent_ex,
                             self.ent_addr, self.ent_dword, self.ent_fword,
                             self.ent_side, self.cb_theme, cb_font,
                             self.ent_editor, self.ent_find]
        # 列幅の保存/復元の対象(設定ファイルのキー -> 一覧)
        self.col_trees = {"cols_doc": self.tree,
                          "cols_exp0": self.panes[0].tree,
                          "cols_exp1": self.panes[1].tree,
                          "cols_find": self.dtree}
        for key, tr in self.col_trees.items():
            apply_col_widths(tr, conf.get(key, ""))
        self.status = tk.StringVar(value="待機中")
        ttk.Label(root, textvariable=self.status, anchor="w",
                  style="Status.TLabel", padding=(10, 5)).pack(fill="x")

        root.bind("<F2>", lambda e: self.do_f2())
        root.bind("<Control-Shift-N>", lambda e: self.new_folder())
        root.bind("<F3>", lambda e: self.focus_word())
        root.bind("<F4>", lambda e: self.focus_widget(self.ent_ex))
        root.bind("<F5>", lambda e: self.do_primary())
        root.bind("<F6>", lambda e: self.cycle_pane())
        root.bind("<Shift-F6>", lambda e: self.cycle_pane(True))
        root.bind("<F7>", lambda e: self.show_fav_menu())
        root.bind("<F8>", lambda e: self.add_favorite())
        root.bind("<F9>", lambda e: self.export_csv())
        root.bind("<Control-d>", lambda e: self.browse())
        root.bind("<Control-l>", lambda e: self.focus_addr())
        root.bind("<Alt-d>", lambda e: self.focus_addr())
        root.bind("<Control-e>", lambda e: self.focus_word())
        root.bind("<Control-f>", lambda e: self.focus_dword())
        root.bind("<Control-r>", lambda e: self.refresh_list())
        root.bind("<Alt-Left>", lambda e: self.go_back())
        root.bind("<Alt-Right>", lambda e: self.go_fwd())
        root.bind("<Alt-Up>", lambda e: self.go_up())
        root.bind("<Alt-Home>", lambda e: self.go_home())
        root.bind("<Control-t>", lambda e: self.new_tab())
        root.bind("<Control-w>", lambda e: self.close_tab_key())
        root.bind("<Control-Tab>", lambda e: self.cycle_tab(1))
        root.bind("<Control-Shift-Tab>", lambda e: self.cycle_tab(-1))
        for d in "123456789":
            root.bind("<Control-Key-%s>" % d,
                      lambda e, dd=d: self.jump_tab_key(dd))
        root.bind("<F1>", lambda e: self.show_help())
        root.bind("<Control-Shift-K>", lambda e: self.save_ext())
        root.bind("<Control-Shift-O>", lambda e: self.save_fav_edit())
        root.bind("<Control-Shift-I>", lambda e: self.full_scan())
        root.bind("<Control-Shift-P>", lambda e: self.stop_scan())
        for i in range(4):
            root.bind("<Alt-Key-%d>" % (i + 1),
                      lambda e, ii=i: self.goto_main_tab(ii))
        root.bind("<Control-g>", lambda e: self.grep_here())
        root.bind("<Control-b>", lambda e: self.focus_side_filter())
        root.bind("<Control-o>", lambda e: self.focus_side_key())
        root.bind("<Control-k>", lambda e: self.focus_list_key())
        root.bind("<Control-Shift-D>", lambda e: self.toggle_two_pane())
        root.bind("<Control-h>", lambda e: self.toggle_showtmp_key())
        root.bind("<Control-Shift-E>", lambda e: self.editor_key())
        root.bind("<Control-Shift-C>", lambda e: self.terminal_key("cmd"))
        root.bind("<Control-Shift-S>",
                  lambda e: self.terminal_key("powershell"))
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.restore_geometry(conf)

        self.apply_theme(theme)
        self.apply_font()
        self.update_scan_label()
        self.build_side()
        self.apply_tab(self.cur_et)
        if self.two_pane and self.panes[0].scope_dir:
            p1 = self.panes[1]
            p1.scope_dir = self.panes[0].scope_dir
            self.start_list(p1.scope_dir, 1, pane=p1)
        self.update_pane_marks()
        self.update_find_info()
        self.root.after(300, self.setup_drop)
        self.root.after(100, self.poll_fq)

    # ---------- ペインの構築/切替 ----------
    def build_pane(self, parent, p):
        p.frame = ttk.Frame(parent)
        p.var_label = tk.StringVar(value="")
        p.label = ttk.Label(p.frame, textvariable=p.var_label,
                            style="PaneHdr.TLabel", anchor="w",
                            padding=(6, 2))
        p.tree = ttk.Treeview(p.frame,
                              columns=("kind", "size", "mtime", "path"),
                              show="tree headings")
        for k in ("name", "kind", "size", "mtime", "path"):
            p.tree.heading(self.colid[k], text=self.head_base[k],
                           command=lambda kk=k, pp=p:
                           self.sort_by_pane(pp, kk))
        # 全列 stretch=False(Explorer同様、列幅は手で決めた値のまま)。
        # 伸縮する列があると、境界をドラッグしてマウスを離した瞬間に
        # Tk(ttk::treeview の drop 処理)がその列へ幅を配り直すため、
        # 名前列の幅は元に戻り、他の列を広げても名前列が縮んで境界が戻ってしまう
        p.tree.column("#0", width=280, minwidth=60, stretch=False)
        p.tree.column("kind", width=100, minwidth=30, anchor="center",
                      stretch=False)
        p.tree.column("size", width=80, minwidth=30, anchor="e",
                      stretch=False)
        p.tree.column("mtime", width=130, minwidth=30, anchor="center",
                      stretch=False)
        p.tree.column("path", width=400, minwidth=30, stretch=False)
        vsb = ttk.Scrollbar(p.frame, orient="vertical",
                            command=p.tree.yview)
        hsb = ttk.Scrollbar(p.frame, orient="horizontal",
                            command=p.tree.xview)
        p.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        p.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="we")
        p.frame.rowconfigure(1, weight=1)
        p.frame.columnconfigure(0, weight=1)
        t = p.tree
        t.bind("<Button-1>", lambda e, pp=p: self.on_tree_click(pp, e))
        t.bind("<Double-1>", self.on_ftree_dblclick)
        t.bind("<Button-3>", self.on_frclick)
        t.bind("<Return>", self.on_ftree_enter)
        t.bind("<Control-Return>", self.on_ftree_place)
        t.bind("<BackSpace>", lambda e: (self.go_back(), "break")[1])
        t.bind("<App>", self.on_fmenu_key)
        t.bind("<Shift-F10>", self.on_fmenu_key)
        t.bind("<Tab>", lambda e: self.on_pane_tab())
        t.bind("<FocusIn>", lambda e, pp=p: self.set_active(pp.idx))
        t.bind("<Control-c>", lambda e: (self.copy_sel(), "break")[1])
        t.bind("<Control-x>", lambda e: (self.cut_sel(), "break")[1])
        t.bind("<Control-v>", lambda e: (self.paste_here(), "break")[1])
        t.bind("<Delete>", lambda e: (self.delete_sel(), "break")[1])
        t.bind("<Shift-Delete>",
               lambda e: (self.delete_sel(True), "break")[1])
        t.bind("<Alt-Return>",
               lambda e: (self.show_properties(), "break")[1])
        t.bind("<KeyPress>", lambda e, pp=p: self.on_tree_keypress(pp, e))
        t.bind("<Control-a>", lambda e: self.select_all())
        t.bind("<<TreeviewSelect>>", lambda e, pp=p: self.on_sel_change(pp))

    def apply_pane_layout(self):
        for p in self.panes:
            try:
                self.mains.forget(p.frame)
            except tk.TclError:
                pass
        self.mains.add(self.panes[0].frame, weight=1)
        if self.two_pane:
            self.mains.add(self.panes[1].frame, weight=1)
        for p in self.panes:
            if self.two_pane:
                p.label.grid(row=0, column=0, columnspan=2, sticky="we")
            else:
                p.label.grid_remove()

    def set_active(self, idx):
        if not self.two_pane:
            idx = 0
        self.cur_pane = idx
        p = self.panes[idx]
        self.var_addr.set(p.scope_dir or "")
        self.var_fword.set(p.word)
        if p.scope_dir is None and p.search:
            self.var_dword.set(p.search[0])
        self.update_headings(p)
        self.update_pane_marks()

    def update_pane_marks(self):
        for i, p in enumerate(self.panes):
            on = (i == self.cur_pane) and self.two_pane
            if p.label is not None:
                p.label.configure(style="PaneHdrOn.TLabel" if on
                                  else "PaneHdr.TLabel")
                if p.scope_dir is None and p.search:
                    txt = "フォルダ名検索: " + p.search[0]
                else:
                    txt = p.scope_dir or "ホーム"
                p.var_label.set(" " + txt)

    def on_tree_click(self, p, event):
        self.set_active(p.idx)

    def on_pane_tab(self):
        if not self.two_pane:
            return None
        self.set_active(1 - self.cur_pane)
        self.focus_flist()
        return "break"

    def toggle_two_pane(self):
        self.two_pane = not self.two_pane
        self.apply_pane_layout()
        if self.two_pane:
            p1 = self.panes[1]
            if p1.scope_dir is None and self.panes[0].scope_dir:
                p1.scope_dir = self.panes[0].scope_dir
            if p1.scope_dir and not p1.scope:
                self.start_list(p1.scope_dir, 1, pane=p1)
            elif p1.scope:
                self.display(pane=p1)
            self.set_active(1)
            self.status.set("2ペイン表示: ON（Tabで行き来できます）")
        else:
            self.set_active(0)
            self.status.set("2ペイン表示: OFF")
        self.focus_flist()
        self.save_conf()
        self.update_pane_marks()

    def reload_pane(self, idx):
        p = self.panes[idx]
        if p.scope_dir:
            # ファイル名検索中なら同じ深さで読み直す(1層だと結果が欠ける)
            self.start_list(p.scope_dir, p.depth, pane=p)
        elif p.search:
            self.run_dsearch(p, *p.search)
        elif idx == self.cur_pane:
            self._home_fill()

    # ---------- ファイル操作(コピー/貼り付け/削除/ドロップ) ----------
    def copy_sel(self):
        p = self.panes[self.cur_pane]
        paths = [p.fpaths[i] for i in p.tree.selection() if i in p.fpaths]
        if not paths:
            self.status.set("コピーする行を選択してください")
            return
        if os.name != "nt":
            self.status.set("ファイルのコピーはWindowsのみ対応です")
            return
        if clip_set_files(paths):
            self.status.set("%d件をクリップボードへコピーしました"
                            "（Ctrl+V / Explorerで貼り付け可）" % len(paths))
        else:
            self.status.set("クリップボードへのコピーに失敗しました")

    def cut_sel(self):
        p = self.panes[self.cur_pane]
        paths = [p.fpaths[i] for i in p.tree.selection() if i in p.fpaths]
        if not paths:
            self.status.set("切り取る行を選択してください")
            return
        if os.name != "nt":
            self.status.set("切り取りはWindowsのみ対応です")
            return
        if clip_set_files(paths, move=True):
            self.status.set("%d件を切り取りました"
                            "（貼り付け先で移動されます）" % len(paths))
        else:
            self.status.set("切り取りに失敗しました")

    def paste_here(self):
        p = self.panes[self.cur_pane]
        if not p.scope_dir:
            self.status.set("貼り付け先のフォルダを開いてから実行してください")
            return
        paths, move = clip_get_files()
        if not paths:
            self.status.set("クリップボードにファイルがありません")
            return
        self.copy_paths_to(paths, p.scope_dir, p.idx, move)

    def delete_sel(self, permanent=False):
        """Delete=ごみ箱へ / Shift+Delete=完全削除(どちらもWindowsの確認付き)"""
        p = self.panes[self.cur_pane]
        paths = [p.fpaths[i] for i in p.tree.selection() if i in p.fpaths]
        if not paths:
            self.status.set("削除する行を選択してください")
            return
        if os.name != "nt":
            self.status.set("削除はWindowsのみ対応です")
            return
        if delete_to_trash(paths, permanent):
            if permanent:
                self.status.set("完全に削除しました: %d件" % len(paths))
            else:
                self.status.set("ごみ箱へ移動しました: %d件" % len(paths))
        else:
            self.status.set("削除をキャンセルしました")
        dirs = {os.path.dirname(sp.rstrip("\\/")) for sp in paths}
        for pp in self.panes:
            if pp.scope_dir == p.scope_dir or pp.scope_dir in dirs:
                self.reload_pane(pp.idx)

    def rename_sel(self):
        p = self.panes[self.cur_pane]
        iid = p.tree.focus()
        path = p.fpaths.get(iid)
        if not path:
            self.status.set("名前を変更する行を選択してください")
            return
        old = os.path.basename(path.rstrip("\\/")) or path
        new = self.ask_text("名前の変更", "新しい名前:", old)
        if not new or new == old:
            return
        if any(ch in new for ch in '\\/:*?"<>|'):
            self.status.set('使えない文字が含まれています: \\ / : * ? " < > |')
            return
        parent = os.path.dirname(path.rstrip("\\/"))
        newpath = os.path.join(parent, new)
        if os.path.exists(newpath):
            self.status.set("同名のファイル/フォルダが既にあります")
            return
        try:
            os.rename(path, newpath)
        except OSError as e:
            self.status.set("変更できません（使用中の可能性）: " + str(e))
            return
        self.status.set("名前を変更しました: " + new)
        old_p = path.rstrip("\\/")
        for pp in self.panes:
            sd = pp.scope_dir
            if sd and (sd == old_p or sd.startswith(old_p + os.sep)):
                pp.scope_dir = newpath + sd[len(old_p):]
                if pp.idx == self.cur_pane:
                    self.var_addr.set(pp.scope_dir)
                    self.sync_tab_name()
            if pp.scope_dir == parent:
                pp.sel_target = newpath
            if pp.scope_dir in (parent, newpath) or pp is p:
                self.reload_pane(pp.idx)
        self.update_pane_marks()

    def new_folder(self):
        if self.cur_tab() != TAB_EXP:
            return
        p = self.panes[self.cur_pane]
        if not p.scope_dir:
            self.status.set("フォルダを開いてから作成してください")
            return
        name = self.ask_text("新しいフォルダー", "フォルダ名:",
                             "新しいフォルダー")
        if not name:
            return
        if any(ch in name for ch in '\\/:*?"<>|'):
            self.status.set('使えない文字が含まれています: \\ / : * ? " < > |')
            return
        path = os.path.join(p.scope_dir, name)
        if os.path.exists(path):
            self.status.set("同名のフォルダ/ファイルが既にあります")
            return
        try:
            os.makedirs(path)
        except OSError as e:
            self.status.set("作成できません: " + str(e))
            return
        self.status.set("フォルダを作成しました: " + name)
        for pp in self.panes:
            if pp.scope_dir == p.scope_dir:
                pp.sel_target = path
                self.reload_pane(pp.idx)

    def do_f2(self):
        if self.cur_tab() == TAB_EXP:
            self.rename_sel()
        else:
            self.focus_widget(self.ent_dir)

    def select_all(self):
        t = self.ftree
        kids = t.get_children()
        if kids:
            t.selection_set(kids)
        return "break"

    def on_sel_change(self, p):
        if p.idx != self.cur_pane:
            return
        n = len(p.tree.selection())
        if n > 1:
            self.status.set("%d件選択中（Ctrl+C / Ctrl+X / Delete で一括操作）"
                            % n)

    def copy_paths_to(self, paths, dest, pidx, move=False):
        srcs = tuple({os.path.dirname(sp.rstrip("\\/")) for sp in paths})

        def run():
            ok = err = 0
            for sp in paths:
                try:
                    src_dir = os.path.dirname(sp.rstrip("\\/"))
                    if move and os.path.abspath(src_dir) == \
                            os.path.abspath(dest):
                        continue
                    if os.path.isdir(sp):
                        low = sp.rstrip("\\/")
                        if dest == low or dest.startswith(low + os.sep):
                            err += 1
                            continue
                    base = os.path.basename(sp.rstrip("\\/")) or "copy"
                    d = unique_dest(os.path.join(dest, base))
                    if move:
                        shutil.move(sp, d)
                    elif os.path.isdir(sp):
                        shutil.copytree(sp, d)
                    else:
                        shutil.copy2(sp, d)
                    ok += 1
                except Exception:
                    err += 1
            self.fq.put(("fopdone", pidx, dest, srcs, ok, err, move))
        self.status.set(("移動中" if move else "コピー中")
                        + "... %d件" % len(paths))
        threading.Thread(target=run, daemon=True).start()

    def on_external_drop(self, idx, paths):
        p = self.panes[idx]
        if not p.scope_dir:
            self.status.set("ドロップ先のフォルダを開いてから"
                            "落としてください")
            return
        self.set_active(idx)
        self.copy_paths_to(paths, p.scope_dir, idx)

    def setup_drop(self):
        if os.name != "nt":
            return
        for p in self.panes:
            try:
                keep = enable_file_drop(
                    p.tree.winfo_id(),
                    lambda paths, ii=p.idx:
                    self.on_external_drop(ii, paths))
                self._dropkeep.append(keep)
            except Exception:
                pass

    # ---------- タブ/キーのディスパッチ ----------
    def cur_tab(self):
        try:
            return self.nb.index("current")
        except tk.TclError:
            return 0

    def on_tab_change(self, event=None):
        tab = self.cur_tab()
        if tab in (TAB_EXP, TAB_FIND):
            self.ensure_index()
        if tab == TAB_FIND:
            if self.find_mode() != "dir":
                self.ensure_findex()
            self.update_find_info()
            self.root.after_idle(self.focus_find_default)

    def do_primary(self):
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.fsearch()
        elif tab == TAB_FIND:
            self.run_find()
        else:
            self.start()

    def focus_word(self):
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.ent_fword.focus_set()
            self.ent_fword.select_range(0, "end")
        elif tab == TAB_FIND:
            self.focus_find_entry()
        else:
            self.focus_widget(self.ent_word)

    def focus_dword(self):
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.ent_dword.focus_set()
            self.ent_dword.select_range(0, "end")
        elif tab == TAB_FIND:
            self.focus_find_entry()
        else:
            self.focus_widget(self.ent_word)

    def focus_addr(self):
        self.nb.select(self.tabf)
        self.ent_addr.focus_set()
        self.ent_addr.select_range(0, "end")

    def refresh_list(self):
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.reload_pane(self.cur_pane)
        elif tab == TAB_FIND:
            self.run_find()

    def focus_side(self):
        kids = self.side.get_children()
        if kids:
            cur = self.side.focus() or kids[0]
            self.side.selection_set(cur)
            self.side.focus(cur)
            self.side.see(cur)
        self.side.focus_set()

    def focus_flist(self):
        t = self.ftree
        kids = t.get_children()
        if kids:
            cur = t.focus() or kids[0]
            t.selection_set(cur)
            t.focus(cur)
            t.see(cur)
        t.focus_set()

    def cycle_pane(self, back=False):
        tab = self.cur_tab()
        if tab == TAB_FIND:
            try:
                cur = self.root.focus_get()
            except (KeyError, tk.TclError):
                cur = None
            if cur is self.dtree:
                self.focus_find_entry()
            else:
                self.focus_find_list()
            return
        if tab != TAB_EXP:
            self.focus_cur_tree()
            return
        order = [self.ent_addr, self.ent_dword, self.ent_fword,
                 self.ent_side, self.side, self.panes[0].tree]
        if self.two_pane:
            order.append(self.panes[1].tree)
        try:
            cur = self.root.focus_get()
        except (KeyError, tk.TclError):
            cur = None
        try:
            i = order.index(cur)
        except ValueError:
            i = -1
        i = (i + (-1 if back else 1)) % len(order)
        w = order[i]
        if w is self.side:
            self.focus_side()
        elif w in self.pane_by_tree:
            self.set_active(self.pane_by_tree[w].idx)
            self.focus_flist()
        else:
            w.focus_set()
            try:
                w.select_range(0, "end")
            except tk.TclError:
                pass

    def focus_cur_tree(self):
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.focus_flist()
        elif tab == TAB_FIND:
            self.focus_find_list()
        else:
            self.focus_tree()

    def goto_main_tab(self, i):
        """Alt+1〜4: 上位タブへ直接ジャンプし、適切な場所にフォーカス"""
        self.nb.select(i)
        if i == TAB_DOC:
            self.ent_word.focus_set()
            self.ent_word.select_range(0, "end")
        elif i == TAB_EXP:
            self.focus_flist()
        elif i == TAB_FIND:
            self.focus_find_entry()
        else:
            self.cb_theme.focus_set()

    def grep_here(self):
        """Ctrl+G: 今いる場所を文書内検索のフォルダにセットして検索へ"""
        tab = self.cur_tab()
        if tab == TAB_EXP:
            iid = self.ftree.focus()
            path = self.fpaths.get(iid)
            if path:
                target = (path if self.fkind.get(iid) == "フォルダ"
                          else os.path.dirname(path))
            else:
                target = self.scope_dir
            if target:
                self.var_dir.set(target)
        elif tab == TAB_FIND:
            path = self.find_focus_place()
            if path:
                self.var_dir.set(path)
        self.nb.select(self.tab1)
        self.ent_word.focus_set()
        self.ent_word.select_range(0, "end")
        self.status.set("文書内検索のフォルダ: "
                        + (self.var_dir.get() or "未設定"))

    def focus_side_key(self):
        if self.cur_tab() == TAB_EXP:
            self.focus_side()

    def focus_list_key(self):
        if self.cur_tab() == TAB_EXP:
            self.focus_flist()

    # ---------- エクスプローラ: タブ ----------
    def save_tab_state(self):
        # 絞り込み語は「検索を実行した語」(p.word)を保存する。入力欄に
        # 打っただけの語を保存すると、戻った時に意図せず絞り込まれるため
        p = self.panes[self.cur_pane]
        t = self.etabs[self.cur_et]
        t.scope_dir = p.scope_dir
        t.scope = p.scope
        t.hist_back = p.hist_back
        t.hist_fwd = p.hist_fwd
        t.sort_key = p.sort_key
        t.sort_desc = p.sort_desc
        t.word = p.word
        t.depth = p.depth
        t.search = p.search

    def apply_tab(self, i):
        self.cur_et = i
        t = self.etabs[i]
        self.invalidate()
        p = self.panes[self.cur_pane]
        p.scope_dir = t.scope_dir
        p.scope = t.scope
        p.hist_back = t.hist_back
        p.hist_fwd = t.hist_fwd
        p.sort_key = t.sort_key
        p.sort_desc = t.sort_desc
        p.word = t.word
        p.depth = t.depth
        p.search = t.search
        self.var_fword.set(t.word)
        self.var_addr.set(t.scope_dir or "")
        self.update_headings(p)
        self.rebuild_tabbar()
        self.update_pane_marks()
        if t.scope_dir is None and not t.search:
            self._home_fill()
        elif t.scope:
            self.display()
        elif t.scope_dir is None:
            self.run_dsearch(p, *t.search)
        else:
            self.start_list(t.scope_dir, t.depth)

    def switch_tab(self, i):
        if i == self.cur_et or not (0 <= i < len(self.etabs)):
            return
        self.save_tab_state()
        self.apply_tab(i)
        self.save_tabs()

    def new_tab(self, path=None):
        self.nb.select(self.tabf)
        self.save_tab_state()
        t = ExpTab()
        if path:
            t.scope_dir = path
            t.name = os.path.basename(path.rstrip("\\/")) or path
        self.etabs.append(t)
        self.apply_tab(len(self.etabs) - 1)
        self.save_tabs()
        if not path:
            self.focus_addr()

    def close_tab_key(self):
        self.nb.select(self.tabf)
        self.close_tab(self.cur_et)

    def close_tab(self, idx):
        if not (0 <= idx < len(self.etabs)):
            return
        if len(self.etabs) == 1:
            self.etabs[0] = ExpTab()
            self.apply_tab(0)
        else:
            closing_cur = (idx == self.cur_et)
            if not closing_cur:
                self.save_tab_state()
            del self.etabs[idx]
            if self.cur_et > idx:
                self.cur_et -= 1
            if closing_cur:
                self.apply_tab(min(idx, len(self.etabs) - 1))
            else:
                self.rebuild_tabbar()
        self.save_tabs()

    def cycle_tab(self, delta):
        self.nb.select(self.tabf)
        self.switch_tab((self.cur_et + delta) % len(self.etabs))

    def jump_tab_key(self, d):
        for i, t in enumerate(self.etabs):
            if t.key == d:
                self.nb.select(self.tabf)
                self.switch_tab(i)
                return

    def rebuild_tabbar(self):
        for w in self.tabbar.winfo_children():
            w.destroy()
        for i, t in enumerate(self.etabs):
            label = t.name
            if t.key:
                label += " (" + t.key + ")"
            if i == self.cur_et:
                label = "● " + label
            b = ttk.Button(self.tabbar, text=label, style="TabBtn.TButton",
                           command=lambda ii=i: self.switch_tab(ii))
            b.pack(side="left", padx=(0, 2))
            b.bind("<Button-3>", lambda e, ii=i: self.tab_menu(e, ii))
            b.bind("<Button-2>", lambda e, ii=i: self.close_tab(ii))
            b.bind("<App>", lambda e, ii=i: self.tab_menu(e, ii))
            b.bind("<Shift-F10>", lambda e, ii=i: self.tab_menu(e, ii))
        plus = ttk.Button(self.tabbar, text="＋", width=3,
                          style="TabBtn.TButton",
                          command=lambda: self.new_tab())
        plus.pack(side="left")
        Tip(plus, "新しいタブ (Ctrl+T)")

    def tab_menu(self, event, idx):
        self.tmenu_idx = idx
        try:
            x = event.x_root
            y = event.y_root
        except AttributeError:
            x = self.tabbar.winfo_rootx()
            y = self.tabbar.winfo_rooty() + 20
        try:
            self.tmenu.tk_popup(x, y)
        finally:
            self.tmenu.grab_release()
        return "break"

    def tab_act(self, act):
        idx = self.tmenu_idx
        if not (0 <= idx < len(self.etabs)):
            return
        t = self.etabs[idx]
        if act == "rename":
            name = self.ask_text("タブ名の変更", "新しいタブ名:", t.name)
            if name:
                t.name = name
                t.custom = True
                self.rebuild_tabbar()
                self.save_tabs()
        elif act == "key":
            v = self.ask_text("ショートカット設定",
                              "Ctrl+ に割り当てる数字 (1〜9、空欄で解除):",
                              t.key)
            if v is None:
                return
            v = v.strip()
            if v == "":
                t.key = ""
            elif v in list("123456789"):
                for o in self.etabs:
                    if o.key == v:
                        o.key = ""
                t.key = v
            else:
                self.status.set("1〜9の数字を1文字で指定してください")
                return
            self.rebuild_tabbar()
            self.save_tabs()
        elif act == "dup":
            self.save_tab_state()
            nt = ExpTab(t.name, t.scope_dir)
            nt.custom = t.custom
            nt.search = t.search
            self.etabs.insert(idx + 1, nt)
            self.apply_tab(idx + 1)
            self.save_tabs()
        elif act == "close":
            self.close_tab(idx)

    def save_tabs(self):
        try:
            with open(TABS_FILE, "w", encoding="utf-8") as f:
                for t in self.etabs:
                    # 検索結果は保存しない(次回起動時はホーム)
                    name = t.name if (t.custom or t.scope_dir) else "ホーム"
                    f.write("\t".join([name, t.scope_dir or "", t.key,
                                       "1" if t.custom else "0"]) + "\n")
        except OSError:
            pass

    def sync_tab_name(self):
        t = self.etabs[self.cur_et]
        p = self.panes[self.cur_pane]
        t.scope_dir = self.scope_dir
        if not t.custom:
            if self.scope_dir:
                t.name = (os.path.basename(self.scope_dir.rstrip("\\/"))
                          or self.scope_dir)
            elif p.search:
                t.name = "検索: " + p.search[0]
            else:
                t.name = "ホーム"
        self.rebuild_tabbar()
        self.save_tabs()

    # ---------- エクスプローラ: ナビゲーション ----------
    def remember_pos(self):
        """今いる場所のカーソル位置を記録(戻ってきた時に復元する)"""
        if self.scope_dir:
            p = self.fpaths.get(self.ftree.focus())
            if p:
                if len(self.sel_memory) > 1000:
                    self.sel_memory.clear()
                self.sel_memory[self.scope_dir] = p

    def invalidate(self, pane=None):
        p = pane if pane is not None else self.panes[self.cur_pane]
        self.gen_seq += 1
        p.gen = self.gen_seq
        p.f_stop.set()

    def _home_fill(self):
        """ホーム: 検索ルート + ユーザーフォルダ + ドライブ(Explorerの「PC」相当)"""
        rows = []
        seen = set()

        def add(name, path):
            key = path.lower().rstrip("\\/")
            if key in seen:
                return
            seen.add(key)
            rows.append(("フォルダ", name, path, name.lower(), -1, 0))

        drives = list_drives()
        dname = {p.lower(): n for n, p in drives}
        for r in load_roots():
            # ルートがドライブ直下(C:\ など)なら「ローカル ディスク (C:)」表記
            add(dname.get(r.lower().rstrip("\\/") + "\\", _basename(r)), r)
        for name, p in user_folders():
            add(name, p)
        for name, p in drives:
            add(name, p)
        self.scope = rows
        self.display("ホーム: 検索ルート・ユーザーフォルダ・ドライブ。"
                     "フォルダ名検索(Ctrl+F)やフォルダ検索タブ(Alt+3)で"
                     "全フォルダから場所を探せます")

    def cur_loc(self):
        """履歴に積む「今の場所」: フォルダのパス / None(ホーム) /
        ("?", 検索語, パス全体?)(フォルダ名検索の結果一覧)"""
        p = self.panes[self.cur_pane]
        if p.scope_dir is None and p.search:
            return ("?",) + tuple(p.search)
        return p.scope_dir

    def _home(self):
        self.remember_pos()
        self.invalidate()
        self.scope_dir = None
        self.panes[self.cur_pane].search = None
        self.var_addr.set("")
        self.panes[self.cur_pane].word = ""
        self.var_fword.set("")
        if self.sort_key == "rank":
            self.sort_key = "name"
        self.sync_tab_name()
        self.update_pane_marks()
        self._home_fill()

    def _goto(self, target):
        self.remember_pos()
        if target is None:
            self._home()
            return
        if isinstance(target, tuple):     # フォルダ名検索の結果へ戻る/進む
            self.show_dsearch(target[1], target[2])
            return
        if not os.path.isdir(target):
            self.status.set("フォルダが存在しません: " + str(target) +
                            "（インデックスが古い可能性）")
            return
        self.scope_dir = target
        self.panes[self.cur_pane].search = None
        self.var_addr.set(target)
        self.panes[self.cur_pane].word = ""
        self.var_fword.set("")
        if self.sort_key == "rank":      # 検索結果の関連度順は通常表示では解除
            self.sort_key = "name"
        self.focus_list = True
        self.sync_tab_name()
        self.update_pane_marks()
        self.start_list(target, 1)

    def navigate(self, path, push=True):
        if not os.path.isdir(path):
            self.status.set("フォルダが存在しません: " + str(path) +
                            "（右クリック→再スキャンで直せます）")
            return
        if push:
            self.hist_back.append(self.cur_loc())
            self.hist_fwd.clear()
        self._goto(path)

    def go_home(self, push=True):
        if push:
            self.hist_back.append(self.cur_loc())
            self.hist_fwd.clear()
        self.focus_list = True
        self.sel_target = self.scope_dir
        self._home()

    def go_back(self):
        if not self.hist_back:
            return
        t = self.hist_back.pop()
        self.hist_fwd.append(self.cur_loc())
        self.focus_list = True
        self.sel_target = self.scope_dir
        self._goto(t)

    def go_fwd(self):
        if not self.hist_fwd:
            return
        t = self.hist_fwd.pop()
        self.hist_back.append(self.cur_loc())
        self.focus_list = True
        self._goto(t)

    def go_up(self):
        if not self.scope_dir:
            return
        parent = os.path.dirname(self.scope_dir.rstrip("\\/"))
        if len(parent) == 2 and parent[1] == ":":
            parent += "\\"     # 「C:」はドライブ直下ではないので C:\ にする
        if parent and parent.rstrip("\\/") != self.scope_dir.rstrip("\\/"):
            self.sel_target = self.scope_dir
            self.navigate(parent)

    def on_addr_enter(self, event=None):
        raw = self.var_addr.get().strip().strip('"')
        if not raw:
            self.go_home()
            return "break"
        # Explorer同様、cmd / powershell / wt と入力するとその場所で端末を開く
        cmd = raw.lower()
        if cmd in ("cmd", "cmd.exe"):
            self.open_terminal_here("cmd")
            self.var_addr.set(self.scope_dir or "")
            return "break"
        if cmd in ("powershell", "powershell.exe", "pwsh", "pwsh.exe"):
            self.open_terminal_here("powershell")
            self.var_addr.set(self.scope_dir or "")
            return "break"
        if cmd in ("wt", "wt.exe"):
            self.open_terminal_here("wt")
            self.var_addr.set(self.scope_dir or "")
            return "break"
        p = os.path.expandvars(os.path.expanduser(raw))
        if not os.path.isabs(p) and not p.startswith("\\\\") \
                and self.scope_dir:
            p = os.path.normpath(os.path.join(self.scope_dir, p))
        if len(p) == 2 and p[1] == ":":
            p += "\\"
        if os.path.isfile(p):
            self.open_file_path(p)
            self.var_addr.set(self.scope_dir or "")
            return "break"
        self.navigate(p)
        return "break"

    def on_entry_escape(self, event=None):
        """検索欄/アドレス欄でEsc: 入力を戻して一覧へフォーカス"""
        if event is not None and event.widget is self.ent_addr:
            self.var_addr.set(self.scope_dir or "")
        self.focus_flist()
        return "break"

    def start_list(self, path, depth, pane=None):
        p = pane if pane is not None else self.panes[self.cur_pane]
        self.gen_seq += 1
        p.gen = self.gen_seq
        p.f_stop.set()
        p.f_stop = threading.Event()
        p.scope = []
        p.depth = depth
        if p.idx == self.cur_pane:
            self.status.set("読み込み中...")
        threading.Thread(target=list_under,
                         args=(path, depth, p.f_stop, self.fq,
                               p.idx, p.gen),
                         daemon=True).start()

    # ---------- エクスプローラ: 検索/表示 ----------
    def ensure_index(self):
        if self.idx_state in ("ready", "loading"):
            return
        if not os.path.exists(DIRIDX_FILE):
            self.status.set("フォルダインデックス未作成: 設定タブでルートを記入し"
                            "「インデックス作成/更新」を実行してください")
            return
        self.idx_state = "loading"
        self.status.set("インデックス読込中...")

        def run():
            self.index.load()
            self.fq.put(("idxdone", len(self.index.paths)))

        threading.Thread(target=run, daemon=True).start()

    def ensure_findex(self):
        """ファイル名インデックスを読み込む(大きいので必要になった時だけ)"""
        if self.fidx_state in ("ready", "loading"):
            return
        if not os.path.exists(FILEIDX_FILE):
            return
        self.fidx_state = "loading"

        def run():
            self.findex.load()
            self.fq.put(("fidxdone", len(self.findex.paths)))

        threading.Thread(target=run, daemon=True).start()

    def update_scan_label(self):
        try:
            t = os.path.getmtime(DIRIDX_FILE)
            self.var_scandate.set("最終スキャン: " + fmt_time(t))
        except OSError:
            self.var_scandate.set("最終スキャン: 未実施")

    def fsearch(self, event=None):
        """F5/検索ボタン: フォーカスと状況から適切な検索を選ぶ"""
        try:
            cur = self.root.focus_get()
        except (KeyError, tk.TclError):
            cur = None
        if cur in (self.ent_dword, self.chk_dpath):
            self.fsearch_folder()
        elif cur in (self.ent_fword, self.dspin):
            self.fsearch_file()
        elif self.scope_dir is None:
            self.fsearch_folder()
        else:
            self.fsearch_file()

    def fsearch_folder(self, event=None):
        """フォルダ名検索: 事前一覧化したインデックス全体から場所を探す"""
        query = self.var_dword.get().strip()
        if not parse_query(query):
            self.status.set("フォルダ名検索: 語を入れてEnter"
                            "（スペース=AND  -語=除外  *?=ワイルドカード）")
            self.focus_dword()
            return "break"
        self.show_dsearch(query, bool(self.var_dpath.get()), push=True)
        return "break"

    def show_dsearch(self, query, full, push=False):
        """フォルダ名検索の結果一覧をアクティブなペインに出す。
        push=True なら今の場所を履歴に積む(BackSpaceで戻れる)"""
        if self.idx_state != "ready":
            # インデックス読込完了後に自動で検索を続行する
            self._pending_fsearch = (query, full, push)
            self.ensure_index()
            if self.idx_state == "loading":
                self.status.set("インデックス読込中... 完了後に検索します")
            return
        p = self.panes[self.cur_pane]
        self.remember_pos()
        loc = self.cur_loc()
        if push and loc != ("?", query, full):
            self.hist_back.append(loc)
            self.hist_fwd.clear()
        self.invalidate()
        self.scope_dir = None
        p.search = (query, full)
        p.word = ""
        self.var_addr.set("")
        self.var_fword.set("")
        self.var_dword.set(query)
        self.sort_key = "rank"
        self.sort_desc = False
        self.focus_list = True
        self.sync_tab_name()
        self.update_pane_marks()
        self.run_dsearch(p, query, full)

    def run_dsearch(self, p, query, full):
        """フォルダ名検索を裏で実行(大きなインデックスでも画面が固まらない)"""
        self.gen_seq += 1
        p.gen = gen = self.gen_seq
        p.f_stop.set()
        p.f_stop = stop = threading.Event()
        p.scope = []
        if p.idx == self.cur_pane:
            self.status.set("フォルダ名検索中...")
        index = self.index

        def work():
            res = index.search(query, MAX_ROWS, full_path=full, stop=stop)
            if res is not None:
                self.fq.put(("dsdone", p.idx, gen, res[0], res[1]))

        threading.Thread(target=work, daemon=True).start()

    def on_showtmp_toggle(self):
        """~$一時ファイルの表示切替: 読み直さずに表示中の一覧を出し直す"""
        self.save_conf()
        for i, p in enumerate(self.panes):
            if i == self.cur_pane or self.two_pane:
                self.display(pane=p)
        self.status.set("一時ファイル(~$...)を"
                        + ("表示します" if self.var_showtmp.get()
                           else "非表示にします") + "（Ctrl+Hで切替）")

    def toggle_showtmp_key(self):
        if self.cur_tab() != TAB_EXP:
            return None
        try:
            cur = self.root.focus_get()
        except (KeyError, tk.TclError):
            cur = None
        if isinstance(cur, (tk.Entry, tk.Spinbox)):
            return None     # 入力欄では Ctrl+H は1文字削除(Tk標準)のまま
        self.var_showtmp.set(not self.var_showtmp.get())
        self.on_showtmp_toggle()
        return "break"

    def on_dpath_toggle(self):
        """「パス」の切替: 検索結果を表示中なら新しい条件で出し直す"""
        self.save_conf()
        p = self.panes[self.cur_pane]
        if p.scope_dir is None and p.search:
            self.show_dsearch(p.search[0], bool(self.var_dpath.get()))

    def fsearch_file(self, event=None):
        """ファイル名検索: 今いる場所から深さN層のファイル名を探す。
        語が空なら通常の一覧(1層)に戻す"""
        if self.scope_dir is None:
            self.status.set("ファイル名検索はフォルダに入ってから。"
                            "場所探しはフォルダ名検索(Ctrl+F)で")
            return "break"
        word = self.var_fword.get().strip()
        self.panes[self.cur_pane].word = word
        if parse_query(word):
            try:
                depth = int(self.var_depth.get())
            except ValueError:
                depth = LIST_DEPTH
            depth = max(1, min(10, depth))
            self.var_depth.set(str(depth))
        else:
            depth = 1
        self.focus_list = True
        self.start_list(self.scope_dir, depth)
        return "break"

    def display(self, note=None, pane=None):
        p = pane if pane is not None else self.panes[self.cur_pane]
        active = (p is self.panes[self.cur_pane])
        cur_path = p.fpaths.get(p.tree.focus())
        toks = parse_query(p.word)
        rows = p.scope
        hidden = 0
        if not self.var_showtmp.get():
            n = len(rows)
            rows = [r for r in rows
                    if not (r[0] == "ファイル" and is_office_temp(r[1]))]
            hidden = n - len(rows)
        if toks:
            # Explorer同様、フォルダ名も対象に含める(r[3]=小文字の名前)
            rows = [r for r in rows if match_query(r[3], toks)]
        rows = self.sort_rows(rows, p)
        if note is None:
            if p.scope_dir is None and p.search:
                what = "パス全体" if p.search[1] else "フォルダ名"
                total = max(p.dtotal, len(rows))
                note = ("%sヒット: %d件（関連度順・列見出しで並べ替え可・"
                        "Enterで中へ・BackSpaceで元の場所へ）"
                        % (what, total))
                if total > len(rows):
                    note += "（上位%d件を表示・語を足して絞込）" % len(rows)
            elif toks:
                note = "名前ヒット: %d 件（深さ%d層まで・フォルダ含む）" % (
                    len(rows), p.depth)
            else:
                note = "%d 件" % len(rows)
            if hidden:
                note += "（~$一時ファイル %d 件は非表示・Ctrl+Hで表示）" % hidden
        self.fill_ftree(p, rows, note if active else "")
        self.update_headings(p)
        kids = p.tree.get_children()
        if kids:
            target = (p.sel_target or cur_path
                      or self.sel_memory.get(p.scope_dir))
            first = kids[0]
            if target:
                hit = False
                for iid, pp in p.fpaths.items():
                    if pp == target:
                        first = iid
                        hit = True
                        break
                if not hit and p.scope_dir is None:
                    tl = target.lower()
                    for iid, pp in p.fpaths.items():
                        pl = pp.lower().rstrip("\\/")
                        if tl == pl or tl.startswith(pl + os.sep):
                            first = iid
                            break
            p.tree.selection_set(first)
            p.tree.focus(first)
            p.tree.see(first)
        p.sel_target = None
        if active and self.focus_list:
            p.tree.focus_set()
            self.focus_list = False

    def sort_rows(self, rows, p):
        if p.sort_key == "rank":
            # フォルダ名検索の結果: 関連度順(7要素目)をそのまま使う
            return sorted(rows, key=lambda r: r[6] if len(r) > 6 else 0)
        if p.sort_key == "kind":
            def val(r):
                if r[0] == "フォルダ":
                    return ""
                return file_type(r[1])[1].lower()
        else:
            idx = {"name": 1, "size": 4,
                   "mtime": 5, "path": 2}[p.sort_key]

            def val(r):
                v = r[idx]
                return v.lower() if isinstance(v, str) else v

        rows = sorted(rows, key=val, reverse=p.sort_desc)
        return sorted(rows, key=lambda r: 0 if r[0] == "フォルダ" else 1)

    def sort_by(self, col):
        p = self.panes[self.cur_pane]
        if p.sort_key == col:
            p.sort_desc = not p.sort_desc
        else:
            p.sort_key = col
            p.sort_desc = False
        self.update_headings(p)
        self.display()

    def sort_by_pane(self, p, col):
        self.set_active(p.idx)
        self.sort_by(col)

    def update_headings(self, pane=None):
        p = pane if pane is not None else self.panes[self.cur_pane]
        for c, base in self.head_base.items():
            mark = ""
            if c == p.sort_key:
                mark = " ▼" if p.sort_desc else " ▲"
            p.tree.heading(self.colid[c], text=base + mark)

    def fill_ftree(self, p, rows, note=""):
        kids = p.tree.get_children()
        if kids:
            p.tree.delete(*kids)
        p.fpaths = {}
        p.fkind = {}
        shown = 0
        for r in rows:
            if shown >= MAX_ROWS:
                break
            kind, name, path, size, mt = r[0], r[1], r[2], r[4], r[5]
            if kind == "フォルダ":
                ikey, label = "folder", "フォルダ"
            else:
                ikey, label = file_type(name)
            stripe = "odd" if shown % 2 else "even"
            tags = ("folder", stripe) if kind == "フォルダ" else (stripe,)
            iid = p.tree.insert(
                "", "end", text=name, image=self.icons[ikey],
                values=(label, fmt_size(size), fmt_time(mt), path),
                tags=tags)
            p.fpaths[iid] = path
            p.fkind[iid] = kind
            shown += 1
        extra = len(rows) - shown
        if extra > 0 and note:
            note += "（表示%d件まで・他%d件は語を足して絞込）" % (MAX_ROWS, extra)
        if note:
            self.status.set(note)

    def poll_fq(self):
        # 1回の処理時間に上限を設け、大量の一覧読込中も画面が固まらないようにする
        deadline = time.time() + 0.08
        busy = False
        try:
            while True:
                if time.time() > deadline:
                    busy = True
                    break
                msg = self.fq.get_nowait()
                k = msg[0]
                if k == "item":
                    _, pidx, g, kind, name, path, size, mt = msg
                    p = self.panes[pidx]
                    if g != p.gen:
                        continue
                    p.scope.append((kind, name, path,
                                    name.lower(), size, mt))
                    if pidx == self.cur_pane and len(p.scope) % 2000 == 0:
                        self.status.set("読み込み中... %d件" % len(p.scope))
                elif k == "listdone":
                    _, pidx, g, n = msg
                    p = self.panes[pidx]
                    if g == p.gen:
                        self.display(pane=p)
                elif k == "fopdone":
                    _, pidx, dest, srcs, ok, err, moved = msg
                    verb = "移動" if moved else "コピー"
                    txt = "%s完了: %d件" % (verb, ok)
                    if err:
                        txt += "（エラー %d件）" % err
                    self.status.set(txt)
                    if moved and ok and os.name == "nt":
                        clip_clear()
                    dirs = set(srcs) | {dest}
                    for pp in self.panes:
                        if pp.scope_dir in dirs:
                            self.reload_pane(pp.idx)
                elif k == "dsdone":
                    _, pidx, g, total, hits = msg
                    p = self.panes[pidx]
                    if g != p.gen or not p.search:
                        continue
                    # 7要素目=関連度順位。sort_key="rank" の間はこの順で表示する
                    p.scope = [("フォルダ", _basename(h), h,
                                _basename(h).lower(), -1, 0, i)
                               for i, h in enumerate(hits)]
                    p.dtotal = total
                    self.display(pane=p)
                elif k == "finddone":
                    _, g, total, hits, secs = msg
                    if g == self._find_gen:
                        self.fill_find(total, hits, secs)
                elif k == "scanprog":
                    self.status.set("スキャン中... %dフォルダ発見" % msg[1])
                elif k == "scandone":
                    self.scanning = False
                    self.idx_state = "ready"
                    txt = "インデックス更新完了: %dフォルダ" % msg[1]
                    if msg[2] is not None:
                        self.fidx_state = "ready"
                        txt += " / %dファイル" % msg[2]
                    self.status.set(txt)
                    self.update_scan_label()
                    self.build_side()
                    self.after_index_change()
                elif k == "scanstop":
                    self.scanning = False
                    self.status.set("スキャン停止（インデックスは未更新）")
                elif k == "idxdone":
                    self.idx_state = "ready"
                    self.status.set("インデックス読込完了: %dフォルダ" % msg[1])
                    pend = self._pending_fsearch
                    self._pending_fsearch = False
                    if pend and self.cur_tab() == TAB_EXP:
                        self.show_dsearch(*pend)
                    self.after_index_change()
                elif k == "fidxdone":
                    self.fidx_state = "ready"
                    self.status.set("ファイル名インデックス読込完了: %dファイル"
                                    % msg[1])
                    self.after_index_change()
        except queue.Empty:
            pass
        self.root.after(10 if busy else 100, self.poll_fq)

    def after_index_change(self):
        """インデックスの読込/更新後: フォルダ検索タブの表示を最新にする"""
        self.update_find_info()
        if parse_query(self.var_find.get()):
            self.run_find()

    def full_scan(self):
        if self.scanning:
            self.status.set("スキャン実行中です")
            return
        self.save_roots()
        roots = load_roots()
        if not roots:
            self.status.set("ルート欄に対象パスを記入してください")
            return
        self.scanning = True
        self.scan_stop = threading.Event()
        self.status.set("スキャン開始...")

        want_files = bool(self.var_fileidx.get())

        def run():
            files = [] if want_files else None
            paths = scan_dirs(roots, self.scan_stop,
                              lambda d, t: self.fq.put(("scanprog", t)),
                              files)
            if self.scan_stop.is_set():
                self.fq.put(("scanstop",))
                return
            save_dirindex(paths)
            self.index.set_paths(paths)
            if files is not None:
                save_dirindex(files, FILEIDX_FILE)
                self.findex.set_paths(files)
            self.fq.put(("scandone", len(paths),
                         len(files) if files is not None else None))

        threading.Thread(target=run, daemon=True).start()

    def rescan_subtree(self, path):
        if self.scanning:
            self.status.set("スキャン実行中です")
            return
        if self.idx_state != "ready":
            self.status.set("先にインデックスを読み込んでください")
            return
        self.scanning = True
        self.scan_stop = threading.Event()
        self.status.set("部分再スキャン中: " + path)

        # ファイル名インデックスがあれば、その部分も作り直す
        want_files = os.path.exists(FILEIDX_FILE)

        def run():
            files = [] if want_files else None
            add = scan_dirs([path], self.scan_stop,
                            lambda d, t: self.fq.put(("scanprog", t)),
                            files)
            if self.scan_stop.is_set():
                self.fq.put(("scanstop",))
                return
            low = path.lower().rstrip("\\/")

            def outside(p):
                pl = p.lower()
                return pl != low and not pl.startswith(low + os.sep)

            allp = [p for p in self.index.paths if outside(p)] + add
            save_dirindex(allp)
            self.index.set_paths(allp)
            nfiles = None
            if files is not None:
                if self.fidx_state != "ready":
                    self.findex.load()
                allf = [p for p in self.findex.paths if outside(p)] + files
                save_dirindex(allf, FILEIDX_FILE)
                self.findex.set_paths(allf)
                nfiles = len(allf)
            self.fq.put(("scandone", len(allp), nfiles))

        threading.Thread(target=run, daemon=True).start()

    def save_roots(self):
        try:
            with open(ROOTS_FILE, "w", encoding="utf-8") as f:
                f.write(self.roots_text.get("1.0", "end").strip() + "\n")
        except OSError as e:
            self.status.set("保存に失敗: " + str(e))
        self.build_side()

    # ---------- エクスプローラ: サイドバー ----------
    def build_side(self):
        self._side_job = None
        open_state = {}
        for top in self.side.get_children():
            for sub in self.side.get_children(top):
                if "grp" in self.side.item(sub, "tags"):
                    open_state[self.side.item(sub, "text")] = \
                        self.side.item(sub, "open")
        for i in self.side.get_children():
            self.side.delete(i)
        words = self.var_sfilter.get().lower().split()
        filtering = bool(words)
        big = len(self.favs) > 30
        fav = self.side.insert("", "end", text="★ お気に入り",
                               open=True, tags=("hdr",))
        groups = {}
        shown = 0
        hidden = 0
        for fi, (name, path) in enumerate(self.favs):
            g, label = fav_group(name)
            if words:
                hay = (g + " " + label + " " + path).lower()
                if not all(w in hay for w in words):
                    continue
            if shown >= 500:
                hidden += 1
                continue
            shown += 1
            parent = fav
            if g:
                if g not in groups:
                    op = True if filtering else open_state.get(g, not big)
                    groups[g] = self.side.insert(
                        fav, "end", text=g, open=op, tags=("grp",))
                parent = groups[g]
            if os.path.splitext(path)[1]:
                ico = self.icons[file_type(path)[0]]
            else:
                ico = self.icons["folder"]
            self.side.insert(parent, "end", text=label,
                             values=(path, str(fi)),
                             image=ico, tags=("itm",))
        if hidden:
            self.side.insert(fav, "end",
                             text="…他 %d 件（語を足して絞り込み）" % hidden,
                             tags=("hdr",))
        rt = self.side.insert("", "end", text="■ 検索ルート",
                              open=True, tags=("hdr",))
        for p in load_roots():
            base = os.path.basename(p.rstrip("\\/")) or p
            if words and not all(w in (base + " " + p).lower()
                                 for w in words):
                continue
            self.side.insert(rt, "end", text=base, values=(p, ""),
                             image=self.icons["folder"], tags=("itm",))
        pc = self.side.insert("", "end", text="■ PC", open=not filtering,
                              tags=("hdr",))
        for name, p in user_folders() + list_drives():
            if words and not all(w in (name + " " + p).lower()
                                 for w in words):
                continue
            self.side.insert(pc, "end", text=name, values=(p, ""),
                             image=self.icons["folder"], tags=("itm",))

    def on_side_filter(self, event=None):
        if event and event.keysym in ("Return", "Escape", "Up", "Down",
                                      "Tab"):
            return
        if self._side_job:
            self.root.after_cancel(self._side_job)
        self._side_job = self.root.after(120, self.build_side)

    def clear_side_filter(self, event=None):
        self.var_sfilter.set("")
        self.build_side()
        self.focus_side()
        return "break"

    def focus_side_filter(self):
        self.nb.select(self.tabf)
        self.ent_side.focus_set()
        self.ent_side.select_range(0, "end")

    def side_act(self, iid):
        tags = self.side.item(iid, "tags")
        if "hdr" in tags or "grp" in tags:
            self.side.item(iid, open=not self.side.item(iid, "open"))
            return
        vals = self.side.item(iid, "values")
        if not vals:
            return
        path = vals[0]
        if os.path.isdir(path):
            self.navigate(path)
        elif os.path.isfile(path):
            self.open_file_path(path)
        else:
            self.status.set("見つかりません: " + path)

    def on_side_press(self, event):
        iid = self.side.identify_row(event.y)
        self._drag_iid = iid
        self._drag_started = False
        self._press_y = event.y
        if iid:
            self.side.selection_set(iid)
            self.side.focus(iid)
        self.side.focus_set()
        return "break"

    def on_side_motion(self, event):
        if not self._drag_iid:
            return
        if not self._drag_started:
            if abs(event.y - self._press_y) < 4:
                return
            vals = self.side.item(self._drag_iid, "values")
            if len(vals) < 2 or vals[1] == "":
                self._drag_iid = None
                return
            self._drag_started = True
            self.side.configure(cursor="hand2")
            self.status.set("移動先へドロップ: "
                            + self.side.item(self._drag_iid, "text"))
        h = self.side.winfo_height()
        if event.y < 16:
            self.side.yview_scroll(-1, "units")
        elif event.y > h - 16:
            self.side.yview_scroll(1, "units")
        tgt = self.side.identify_row(event.y)
        if tgt and tgt != self._drag_iid:
            self.side.selection_set(tgt)

    def on_side_release(self, event):
        iid = self._drag_iid
        started = self._drag_started
        self._drag_iid = None
        self._drag_started = False
        self.side.configure(cursor="")
        if not iid:
            return
        if not started:
            self.side_act(iid)
            return
        tgt = self.side.identify_row(event.y)
        if not tgt or tgt == iid:
            self.build_side()
            return
        self.move_fav(iid, tgt, event.y)

    def move_fav(self, src_iid, tgt_iid, y):
        vals = self.side.item(src_iid, "values")
        try:
            si = int(vals[1])
        except (IndexError, ValueError):
            return
        if not (0 <= si < len(self.favs)):
            return
        label = fav_group(self.favs[si][0])[1]
        path = self.favs[si][1]
        ttags = self.side.item(tgt_iid, "tags")
        tvals = self.side.item(tgt_iid, "values")
        ttext = self.side.item(tgt_iid, "text")
        orig = self.favs.pop(si)
        if "grp" in ttags:
            g = ttext
            idxs = [i for i, (n, _p) in enumerate(self.favs)
                    if fav_group(n)[0] == g]
            pos = (idxs[-1] + 1) if idxs else len(self.favs)
        elif "itm" in ttags and len(tvals) >= 2 and tvals[1] != "":
            ti = int(tvals[1])
            if ti > si:
                ti -= 1
            g = fav_group(self.favs[ti][0])[0]
            bbox = self.side.bbox(tgt_iid)
            after = bool(bbox) and (y > bbox[1] + bbox[3] // 2)
            pos = ti + (1 if after else 0)
        elif "hdr" in ttags and ttext.startswith("★"):
            g = ""
            pos = len(self.favs)
        else:
            self.favs.insert(si, orig)
            self.build_side()
            return
        newname = (g + "/" + label) if g else label
        self.favs.insert(pos, (newname, path))
        save_favs(self.favs)
        self.refresh_fav_text()
        self.build_side()
        self.select_fav(pos)
        self.status.set("お気に入りを移動しました: " + label)

    def select_fav(self, pos):
        def walk(parent):
            for c in self.side.get_children(parent):
                v = self.side.item(c, "values")
                if len(v) >= 2 and v[1] == str(pos):
                    self.side.selection_set(c)
                    self.side.focus(c)
                    self.side.see(c)
                    return True
                if walk(c):
                    return True
            return False
        walk("")

    def on_side_move_key(self, delta):
        iid = self.side.focus()
        vals = self.side.item(iid, "values") if iid else ()
        if len(vals) < 2 or vals[1] == "":
            return "break"
        i = int(vals[1])
        j = i + delta
        if not (0 <= j < len(self.favs)):
            return "break"
        self.favs[i], self.favs[j] = self.favs[j], self.favs[i]
        save_favs(self.favs)
        self.refresh_fav_text()
        self.build_side()
        self.select_fav(j)
        return "break"

    def on_side_enter(self, event):
        iid = self.side.focus()
        if iid:
            self.side_act(iid)
        return "break"

    # ---------- エクスプローラ: 一覧の操作 ----------
    def on_ftree_enter(self, event=None):
        iid = self.ftree.focus()
        path = self.fpaths.get(iid)
        if not path:
            return "break"
        if self.fkind.get(iid) == "フォルダ":
            self.navigate(path)
        else:
            self.open_file_path(path)
        return "break"

    def on_ftree_dblclick(self, event):
        """ダブルクリック: 行の上なら開く。列見出しの境界なら列幅を内容に
        合わせる(Explorer同様)。見出し上のダブルクリックでは何も開かない"""
        t = event.widget
        region = t.identify_region(event.x, event.y)
        if region == "separator":
            autofit_column(t, t.identify_column(event.x), self.cur_font)
            return "break"
        if region not in ("tree", "cell"):
            return "break"
        return self.on_ftree_enter()

    def on_ftree_place(self, event=None):
        iid = self.ftree.focus()
        path = self.fpaths.get(iid)
        if not path:
            return "break"
        if self.fkind.get(iid) == "フォルダ":
            self.open_folder_plain(path)
        else:
            self.open_folder_sel(path)
        return "break"

    def on_frclick(self, event):
        p = self.pane_by_tree.get(event.widget)
        if p is not None:
            self.set_active(p.idx)
        iid = self.ftree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.ftree.selection():
            self.ftree.selection_set(iid)
        self.ftree.focus(iid)
        self.fmenu_iid = iid
        try:
            self.fmenu.tk_popup(event.x_root, event.y_root)
        finally:
            self.fmenu.grab_release()

    def on_fmenu_key(self, event):
        iid = self.ftree.focus()
        if not iid:
            return "break"
        self.fmenu_iid = iid
        bbox = self.ftree.bbox(iid)
        if bbox:
            x = self.ftree.winfo_rootx() + bbox[0] + 60
            y = self.ftree.winfo_rooty() + bbox[1] + bbox[3]
        else:
            x = self.ftree.winfo_rootx() + 60
            y = self.ftree.winfo_rooty() + 20
        try:
            self.fmenu.tk_popup(x, y)
        finally:
            self.fmenu.grab_release()
        return "break"

    def fm_do(self, act):
        iid = self.fmenu_iid
        path = self.fpaths.get(iid)
        if not path:
            return
        isdir = self.fkind.get(iid) == "フォルダ"
        if act == "open":
            if isdir:
                self.navigate(path)
            else:
                self.open_file_path(path)
        elif act == "place":
            if isdir:
                self.open_folder_plain(path)
            else:
                self.open_folder_sel(path)
        elif act == "newtab":
            self.new_tab(path if isdir else os.path.dirname(path))
        elif act == "fav":
            self._add_fav(path)
        elif act == "todoc":
            self.var_dir.set(path if isdir else os.path.dirname(path))
            self.nb.select(self.tab1)
            self.focus_widget(self.ent_word)
            self.status.set("文書内検索のフォルダに設定しました")
        elif act == "rescan":
            if isdir:
                self.rescan_subtree(path)
            else:
                self.status.set("フォルダ行を選んで実行してください")
        elif act == "cp_full":
            self.to_clip(path)
        elif act == "cp_dir":
            self.to_clip(path if isdir else os.path.dirname(path))
        elif act == "cp_name":
            self.to_clip(os.path.basename(path.rstrip("\\/")))
        elif act == "editor":
            sel = self.sel_paths()
            self.open_in_editor(sel if path in sel else [path])
        elif act == "openas":
            if isdir:
                self.status.set("ファイル行で実行してください")
            elif not shell_verb(path, "openas"):
                self.status.set("「プログラムから開く」を表示できません")
        elif act.startswith("term_"):
            self.open_terminal_here(
                act[5:], path if isdir else os.path.dirname(path))
        elif act == "props":
            self.show_properties(path)

    def to_clip(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.set("コピーしました: " + text)

    # ---------- フォルダ検索タブ(Everything風) ----------
    def build_find_tab(self, conf):
        """入力するそばからインデックス全体のフォルダ名を検索するタブ。
        戻り値は区切り線(テーマ色の対象)"""
        tab = self.tabd
        top = ttk.Frame(tab, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Label(top, text="名前:").pack(side="left")
        self.var_find = tk.StringVar()
        self.ent_find = ttk.Entry(top, textvariable=self.var_find)
        self.ent_find.pack(side="left", fill="x", expand=True, padx=(6, 0))
        Tip(self.ent_find, "入力するそばからフォルダ/ファイル名を検索 (Ctrl+F)\n"
                           "スペース=AND  -語=除外  *?=ワイルドカード\n"
                           "\\ や / を含む語はパス全体で照合  "
                           "↓で結果へ  Escで消去")
        self.var_find.trace_add("write", lambda *a: self.schedule_find())
        self.ent_find.bind("<Return>", self.on_find_enter)
        self.ent_find.bind("<Down>", self.on_find_down)
        self.ent_find.bind("<Escape>", self.on_find_escape)
        ttk.Label(top, text="  対象:").pack(side="left")
        mode = conf.get("fmode", "dir")
        self.var_fmode = tk.StringVar(
            value=FIND_MODES.get(mode, FIND_MODES["dir"]))
        cb = ttk.Combobox(top, textvariable=self.var_fmode, width=8,
                          state="readonly", values=list(FIND_MODES.values()))
        cb.pack(side="left", padx=(4, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self.on_fmode_change())
        self.cb_fmode = cb
        Tip(cb, "フォルダ / ファイル / 両方 (Ctrl+Shift+F で切替)\n"
                "ファイルを探すには設定タブの「ファイル名もインデックスに"
                "含める」をONにしてインデックス更新")
        self.var_fpath = tk.BooleanVar(value=conf.get("fpath", "0") == "1")
        chk = ttk.Checkbutton(top, text="パス全体も対象",
                              variable=self.var_fpath,
                              command=self.on_fpath_toggle)
        chk.pack(side="left", padx=(10, 0))
        Tip(chk, "ON: パスのどこかに含まれれば一致\n"
                 "OFF: 名前そのものに含まれる場合だけ一致")
        b = ttk.Button(top, text="インデックス更新", takefocus=False,
                       command=self.full_scan)
        b.pack(side="left", padx=(10, 0))
        Tip(b, "設定タブの検索ルートを再スキャン (Ctrl+Shift+I)")

        self.var_find_info = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.var_find_info, style="Sub.TLabel",
                  padding=(12, 0, 10, 4)).pack(fill="x")
        gold = tk.Frame(tab, height=1)
        gold.pack(fill="x", padx=10)

        mid = ttk.Frame(tab)
        mid.pack(fill="both", expand=True, padx=10, pady=(6, 8))
        t = ttk.Treeview(mid, columns=("place",), show="tree headings")
        self.dtree = t
        t.heading("#0", command=lambda: self.find_sort_by("name"))
        t.heading("place", command=lambda: self.find_sort_by("path"))
        # 名前列は stretch=False(手で決めた幅を保つ)。余白は場所列で吸収
        t.column("#0", width=300, minwidth=60, stretch=False)
        t.column("place", width=620, minwidth=150, stretch=True)
        self.update_find_headings()
        vsb = ttk.Scrollbar(mid, orient="vertical", command=t.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        t.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="we")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        t.bind("<Double-1>", self.on_find_dblclick)
        t.bind("<Return>", lambda e: self.find_open("app"))
        t.bind("<Shift-Return>", lambda e: self.find_open("newtab"))
        t.bind("<Control-Return>", lambda e: self.find_open("explorer"))
        t.bind("<Alt-Return>", lambda e: self.find_open("props"))
        t.bind("<Button-3>", self.on_find_rclick)
        t.bind("<App>", self.on_find_menu_key)
        t.bind("<Shift-F10>", self.on_find_menu_key)
        t.bind("<Escape>", lambda e: (self.focus_find_entry(), "break")[1])
        t.bind("<Control-c>", lambda e: (self.find_copy(), "break")[1])
        t.bind("<BackSpace>", self.on_find_backspace)
        for w in (t, self.ent_find):
            w.bind("<Control-Shift-F>",
                   lambda e: (self.cycle_fmode(), "break")[1])
        t.bind("<KeyPress>", self.on_find_tree_key)

        m = tk.Menu(self.root, tearoff=0)
        self.dmenu = m
        m.add_command(label="開く (Enter)",
                      command=lambda: self.fd_do("app"))
        m.add_command(label="新しいタブで開く (Shift+Enter)",
                      command=lambda: self.fd_do("newtab"))
        m.add_command(label="場所(親フォルダ)を開いて選択",
                      command=lambda: self.fd_do("parent"))
        m.add_command(label="Explorerで開く (Ctrl+Enter)",
                      command=lambda: self.fd_do("explorer"))
        m.add_command(label="サクラエディタで開く (Ctrl+Shift+E)",
                      command=lambda: self.fd_do("editor"))
        m.add_separator()
        m.add_command(label="この場所で文書内検索 (Ctrl+G)",
                      command=lambda: self.fd_do("todoc"))
        m.add_command(label="お気に入りに追加 (F8)",
                      command=lambda: self.fd_do("fav"))
        self.dmenu_term = tk.Menu(m, tearoff=0)
        self.dmenu_term.add_command(
            label="コマンドプロンプト (Ctrl+Shift+C)",
            command=lambda: self.fd_do("term_cmd"))
        self.dmenu_term.add_command(
            label="PowerShell (Ctrl+Shift+S)",
            command=lambda: self.fd_do("term_powershell"))
        self.dmenu_term.add_command(
            label="Windows Terminal", command=lambda: self.fd_do("term_wt"))
        m.add_cascade(label="この場所で端末を開く", menu=self.dmenu_term)
        m.add_command(label="プロパティ (Alt+Enter)",
                      command=lambda: self.fd_do("props"))
        m.add_separator()
        m.add_command(label="フルパスをコピー (Ctrl+C)",
                      command=lambda: self.fd_do("cp_full"))
        m.add_command(label="名前をコピー",
                      command=lambda: self.fd_do("cp_name"))
        m.add_separator()
        m.add_command(label="このフォルダ以下を再スキャン",
                      command=lambda: self.fd_do("rescan"))
        return gold

    def find_mode(self):
        """フォルダ検索タブの対象: "dir" / "file" / "both" """
        v = self.var_fmode.get()
        for k, label in FIND_MODES.items():
            if label == v:
                return k
        return "dir"

    def on_fmode_change(self):
        self.save_conf()
        if self.find_mode() != "dir":
            self.ensure_findex()
        self.update_find_info()
        self.run_find()

    def cycle_fmode(self):
        keys = list(FIND_MODES)
        k = keys[(keys.index(self.find_mode()) + 1) % len(keys)]
        self.var_fmode.set(FIND_MODES[k])
        self.on_fmode_change()
        self.status.set("フォルダ検索の対象: " + FIND_MODES[k])

    def find_state(self):
        return (self.var_find.get(), bool(self.var_fpath.get()),
                self.find_sort, self.find_mode())

    def update_find_info(self, text=None):
        """フォルダ検索タブの案内行(件数やインデックスの状態)を更新"""
        if text is None:
            if not parse_query(self.var_find.get()):
                text = self.find_idle_text()
            else:
                return
        self.var_find_info.set(text)

    def find_idle_text(self):
        mode = self.find_mode()
        if mode != "dir" and not os.path.exists(FILEIDX_FILE):
            return ("ファイル名インデックスがありません: 設定タブの「ファイル名も"
                    "インデックスに含める」をONにして「インデックス更新」"
                    "(Ctrl+Shift+I) を押してください")
        if self.idx_state == "ready" and (
                mode == "dir" or self.fidx_state == "ready"):
            try:
                when = fmt_time(os.path.getmtime(DIRIDX_FILE))
            except OSError:
                when = "未保存"
            what = []
            if mode != "file":
                what.append("%s フォルダ" % format(len(self.index.paths), ","))
            if mode != "dir":
                nf = len(self.findex.paths)
                what.append("%s ファイル" % format(nf, ","))
            return ("全 %s から検索します（最終スキャン: %s）。"
                    "スペース=AND  -語=除外  *?=ワイルドカード"
                    % (" / ".join(what), when))
        if "loading" in (self.idx_state, self.fidx_state):
            return "インデックス読込中..."
        if not os.path.exists(DIRIDX_FILE):
            return ("インデックス未作成: 設定タブの「検索ルート」に対象"
                    "（例: C:\\ や \\\\server\\share）を記入し、"
                    "「インデックス更新」(Ctrl+Shift+I) を押してください")
        return "インデックス未読込（このタブを開くと読み込みます）"

    def schedule_find(self, delay=FIND_DELAY_MS):
        """入力のたびに呼ばれる。打ち終わるのを少し待ってから検索する"""
        if self._find_job:
            self.root.after_cancel(self._find_job)
        self._find_job = self.root.after(delay, self.run_find)

    def run_find(self):
        if self._find_job:
            self.root.after_cancel(self._find_job)
            self._find_job = None
        self._find_stop.set()           # 前の検索が走っていれば打ち切る
        self._find_gen += 1
        gen = self._find_gen
        state = self.find_state()
        query, full, (sort, desc), mode = state
        if not parse_query(query):
            self._find_focus_after = False
            self.fill_find(0, [], None)
            return
        # 対象に応じたインデックス: (インデックス, ファイルか)
        targets = []
        if mode != "file":
            if self.idx_state != "ready":
                self.ensure_index()
            targets.append((self.index, False))
        if mode != "dir":
            if self.fidx_state != "ready":
                self.ensure_findex()
            targets.append((self.findex, True))
        if (mode != "file" and self.idx_state != "ready") or \
                (mode != "dir" and self.fidx_state != "ready"):
            self.update_find_info(self.find_idle_text())
            return
        stop = threading.Event()
        self._find_stop = stop
        self._find_running = state
        t0 = time.time()

        def work():
            total = 0
            lists = []
            for index, isfile in targets:
                res = index.search(query, MAX_ROWS, full_path=full,
                                   sort=sort, desc=desc, stop=stop,
                                   with_keys=True)
                if res is None:
                    return
                total += res[0]
                lists.append([(k, p, isfile) for k, p in res[1]])
            # 各インデックスの上位を同じ並べ替えキーで混ぜる
            merged = heapq.merge(*lists, key=lambda x: x[0], reverse=desc)
            hits = [(p, isfile) for _k, p, isfile in merged][:MAX_ROWS]
            self.fq.put(("finddone", gen, total, hits, time.time() - t0))

        threading.Thread(target=work, daemon=True).start()

    def fill_find(self, total, hits, secs):
        t = self.dtree
        keep = self.find_paths.get(t.focus())
        kids = t.get_children()
        if kids:
            t.delete(*kids)
        self.find_paths = {}
        self.find_isfile = {}
        sel = None
        for i, (p, isfile) in enumerate(hits):
            name = _basename(p)
            stripe = "odd" if i % 2 else "even"
            if isfile:
                ico, tags = self.icons[file_type(name)[0]], (stripe,)
            else:
                ico, tags = self.icons["folder"], ("folder", stripe)
            iid = t.insert("", "end", text=name, image=ico,
                           values=(parent_of(p),), tags=tags)
            self.find_paths[iid] = p
            self.find_isfile[iid] = isfile
            if p == keep:
                sel = iid
        kids = t.get_children()
        if kids:
            cur = sel or kids[0]
            t.selection_set(cur)
            t.focus(cur)
            t.see(cur)
        if secs is None:
            self._find_shown = None
            self.update_find_info()
            return
        self._find_shown = self._find_running
        sort, desc = self.find_sort
        order = {"rank": "関連度順", "name": "名前順",
                 "path": "場所順"}[sort] + (" (降順)" if desc else "")
        txt = "%s 件（%.2f秒・%s）" % (format(total, ","), secs, order)
        if total > len(hits):
            txt += "  上位 %s 件を表示中。語を足すと絞り込めます" % format(
                len(hits), ",")
        elif not total:
            txt = "該当なし（%.2f秒）。インデックスが古い場合は「インデックス更新」" \
                  % secs
        self.update_find_info(txt)
        if self.cur_tab() == TAB_FIND:
            self.status.set("フォルダ検索: " + txt)
        if self._find_focus_after:
            self._find_focus_after = False
            self.focus_find_list()

    def update_find_headings(self):
        sort, desc = self.find_sort
        for col, key, base in (("#0", "name", "名前"),
                               ("place", "path", "場所")):
            mark = (" ▼" if desc else " ▲") if sort == key else ""
            self.dtree.heading(col, text=base + mark)

    def find_sort_by(self, key):
        """列見出しクリック: 昇順 → 降順 → 関連度順 の順に切り替え"""
        sort, desc = self.find_sort
        if sort != key:
            self.find_sort = (key, False)
        elif not desc:
            self.find_sort = (key, True)
        else:
            self.find_sort = ("rank", False)
        self.update_find_headings()
        self.run_find()

    def on_fpath_toggle(self):
        self.save_conf()
        self.run_find()

    def focus_find_entry(self):
        self.nb.select(self.tabd)
        self.ent_find.focus_set()
        self.ent_find.select_range(0, "end")
        self.ent_find.icursor("end")

    def focus_find_list(self):
        self.nb.select(self.tabd)
        t = self.dtree
        kids = t.get_children()
        if not kids:
            self.focus_find_entry()
            return
        cur = t.focus() or kids[0]
        t.selection_set(cur)
        t.focus(cur)
        t.see(cur)
        t.focus_set()

    def focus_find_default(self):
        """タブを開いた時: 一覧か検索欄にフォーカスがなければ検索欄へ"""
        try:
            cur = self.root.focus_get()
        except (KeyError, tk.TclError):
            cur = None
        if self.cur_tab() == TAB_FIND and cur not in (self.dtree,
                                                      self.ent_find):
            self.focus_find_entry()

    def on_find_enter(self, event=None):
        """検索欄でEnter: 結果が出ていれば一覧へ(出る前なら出た時に)"""
        if self._find_job is None and self._find_shown == self.find_state():
            self.focus_find_list()
        else:
            self._find_focus_after = True
            self.run_find()
        return "break"

    def on_find_down(self, event=None):
        if self.dtree.get_children():
            self.focus_find_list()
        return "break"

    def on_find_escape(self, event=None):
        if self.var_find.get():
            self.var_find.set("")
        return "break"

    def on_find_tree_key(self, event):
        """一覧上で文字を打ったら検索欄へ送る(Everything同様)"""
        ch = event.char
        if not ch or len(ch) != 1 or not ch.isprintable():
            return None
        if event.state & 0x4 or event.state & 0x20000:   # Ctrl / Alt
            return None
        self.ent_find.focus_set()
        self.ent_find.select_clear()
        self.ent_find.insert("end", ch)
        self.ent_find.icursor("end")
        return "break"

    def on_find_backspace(self, event=None):
        s = self.var_find.get()
        if s:
            self.var_find.set(s[:-1])
        self.ent_find.focus_set()
        self.ent_find.icursor("end")
        return "break"

    def find_focus_path(self):
        return self.find_paths.get(self.dtree.focus())

    def find_focus_place(self):
        """フォーカス行のフォルダ(ファイル行ならその親フォルダ)"""
        iid = self.dtree.focus()
        path = self.find_paths.get(iid)
        if path and self.find_isfile.get(iid):
            return parent_of(path)
        return path

    def find_copy(self):
        paths = [self.find_paths[i] for i in self.dtree.selection()
                 if i in self.find_paths]
        if paths:
            self.to_clip("\n".join(paths))

    def on_find_dblclick(self, event):
        t = self.dtree
        region = t.identify_region(event.x, event.y)
        if region == "separator":
            autofit_column(t, t.identify_column(event.x), self.cur_font)
        elif region in ("tree", "cell"):
            self.find_open("app")
        return "break"

    def on_find_rclick(self, event):
        t = self.dtree
        iid = t.identify_row(event.y)
        if not iid:
            return
        if iid not in t.selection():
            t.selection_set(iid)
        t.focus(iid)
        self.dmenu_iid = iid
        try:
            self.dmenu.tk_popup(event.x_root, event.y_root)
        finally:
            self.dmenu.grab_release()

    def on_find_menu_key(self, event=None):
        t = self.dtree
        iid = t.focus()
        if not iid:
            return "break"
        self.dmenu_iid = iid
        bbox = t.bbox(iid)
        if bbox:
            x = t.winfo_rootx() + bbox[0] + 60
            y = t.winfo_rooty() + bbox[1] + bbox[3]
        else:
            x = t.winfo_rootx() + 60
            y = t.winfo_rooty() + 20
        try:
            self.dmenu.tk_popup(x, y)
        finally:
            self.dmenu.grab_release()
        return "break"

    def find_open(self, how):
        self.dmenu_iid = self.dtree.focus()
        self.fd_do(how)
        return "break"

    def fd_do(self, act):
        """フォルダ検索の結果行(フォルダまたはファイル)に対する操作"""
        path = self.find_paths.get(self.dmenu_iid)
        if not path:
            return
        isfile = self.find_isfile.get(self.dmenu_iid, False)
        # ファイルなら「その場所」= 親フォルダを対象にする操作
        place = parent_of(path) if isfile else path
        if act in ("app", "newtab", "parent", "explorer", "todoc", "fav",
                   "props", "editor") or act.startswith("term_"):
            ok = os.path.isfile(path) if isfile else os.path.isdir(path)
            if not ok:
                self.status.set(("ファイル" if isfile else "フォルダ")
                                + "が存在しません: " + path +
                                "（インデックスが古い可能性。右クリック→"
                                "再スキャン、またはインデックス更新）")
                return
        if act == "app":
            if isfile:
                self.open_file_path(path)
            else:
                self.open_dir_in_app(path)
        elif act == "newtab":
            self.open_dir_in_app(place, newtab=True,
                                 select=path if isfile else None)
        elif act == "parent":
            parent = parent_of(path)
            if parent:
                self.open_dir_in_app(parent, select=path)
        elif act == "explorer":
            if isfile:
                self.open_folder_sel(path)
            else:
                self.open_folder_plain(path)
        elif act == "editor":
            if isfile:
                sel = [self.find_paths[i] for i in self.dtree.selection()
                       if self.find_isfile.get(i)]
                self.open_in_editor(sel if path in sel else [path])
            else:
                self.status.set("ファイル行で実行してください")
        elif act == "todoc":
            self.var_dir.set(place)
            self.nb.select(self.tab1)
            self.focus_widget(self.ent_word)
            self.status.set("文書内検索のフォルダに設定しました")
        elif act == "fav":
            self._add_fav(path)
        elif act.startswith("term_"):
            self.open_terminal_here(act[5:], place)
        elif act == "props":
            self.show_properties(path)
        elif act == "cp_full":
            if len(self.dtree.selection()) > 1:
                self.find_copy()
            else:
                self.to_clip(path)
        elif act == "cp_name":
            self.to_clip(_basename(path))
        elif act == "rescan":
            self.rescan_subtree(place)

    def open_dir_in_app(self, path, newtab=False, select=None):
        """フォルダをこのアプリのエクスプローラタブで開く"""
        self.nb.select(self.tabf)
        self.focus_list = True
        if newtab:
            self.new_tab(path)
            if select:
                self.sel_target = select
        else:
            if select:
                self.sel_target = select
            self.navigate(path)

    # ---------- エディタ/端末/シェル連携 ----------
    def sel_paths(self):
        """アクティブなペインで選択中の行のパス一覧"""
        p = self.panes[self.cur_pane]
        return [p.fpaths[i] for i in p.tree.selection() if i in p.fpaths]

    def editor_exe(self):
        exe = find_editor(self.var_editor.get())
        if not exe:
            self.status.set("サクラエディタが見つかりません。設定タブの"
                            "「テキストエディタ」に sakura.exe のパスを"
                            "指定してください")
        return exe

    def open_in_editor(self, paths):
        """サクラエディタ(設定で変更可)でファイルを開く。フォルダは対象外"""
        files = [p for p in paths if p and os.path.isfile(p)]
        if not files:
            self.status.set("ファイル行を選んで実行してください"
                            "（フォルダはエディタで開けません）")
            return
        exe = self.editor_exe()
        if not exe:
            return
        files = files[:10]
        try:
            for f in files:
                subprocess.Popen([exe, f])
        except OSError as e:
            self.status.set("エディタを起動できません: " + str(e))
            return
        self.status.set("%s で開きました: %d件" % (
            os.path.basename(exe), len(files)))

    def editor_key(self):
        """Ctrl+Shift+E: 表示中のタブに応じて対象ファイルを決める"""
        tab = self.cur_tab()
        if tab == TAB_EXP:
            paths = self.sel_paths()
            if not paths:
                paths = [self.fpaths.get(self.ftree.focus(), "")]
        elif tab == TAB_DOC:
            paths = [self.paths.get(self.tree.focus(), "")]
        elif tab == TAB_FIND:
            self.find_open("editor")
            return
        else:
            return
        self.open_in_editor(paths)

    def doc_dir_of(self, iid):
        """文書内検索の行 -> そのファイルのフォルダ(なければ検索フォルダ)"""
        path = self.paths.get(iid, "")
        if path:
            return os.path.dirname(path)
        d = self.var_dir.get().strip().strip('"')
        return d if os.path.isdir(d) else None

    def cur_dir_for_shell(self):
        """端末/プロパティの既定対象: 今いる場所。ホームや検索結果上なら
        フォーカス行のフォルダ(ファイルならその親)"""
        if self.scope_dir:
            return self.scope_dir
        iid = self.ftree.focus()
        path = self.fpaths.get(iid)
        if not path:
            return None
        return path if self.fkind.get(iid) == "フォルダ" else \
            os.path.dirname(path)

    def open_terminal_here(self, kind, path=None):
        if path is None:
            path = self.cur_dir_for_shell()
        if not path or not os.path.isdir(path):
            self.status.set("フォルダを開いてから実行してください")
            return
        label = {"cmd": "コマンドプロンプト", "powershell": "PowerShell",
                 "wt": "Windows Terminal"}.get(kind, kind)
        if open_terminal(kind, path):
            self.status.set("%s を開きました: %s" % (label, path))
        elif kind == "wt":
            self.status.set("Windows Terminal (wt.exe) が見つかりません")
        else:
            self.status.set(label + " を起動できませんでした")

    def terminal_key(self, kind):
        """Ctrl+Shift+C / S: 表示中のタブに応じた場所で端末を開く"""
        tab = self.cur_tab()
        if tab == TAB_EXP:
            self.open_terminal_here(kind)
        elif tab == TAB_DOC:
            self.open_terminal_here(kind, self.doc_dir_of(self.tree.focus()))
        elif tab == TAB_FIND:
            self.open_terminal_here(kind, self.find_focus_place())

    def show_properties(self, path=None):
        if path is None:
            path = self.fpaths.get(self.ftree.focus()) or self.scope_dir
        if not path:
            self.status.set("行を選ぶかフォルダを開いてから実行してください")
            return
        if os.name != "nt" or not shell_verb(path, "properties"):
            self.status.set("プロパティを表示できません: " + path)

    def new_text_file(self):
        if self.cur_tab() != TAB_EXP:
            return
        p = self.panes[self.cur_pane]
        if not p.scope_dir:
            self.status.set("フォルダを開いてから作成してください")
            return
        name = self.ask_text("新しいテキスト ファイル", "ファイル名:",
                             "新しいテキスト ドキュメント.txt")
        if not name:
            return
        if any(ch in name for ch in '\\/:*?"<>|'):
            self.status.set('使えない文字が含まれています: \\ / : * ? " < > |')
            return
        path = os.path.join(p.scope_dir, name)
        if os.path.exists(path):
            self.status.set("同名のファイル/フォルダが既にあります")
            return
        try:
            with open(path, "x", encoding="utf-8"):
                pass
        except OSError as e:
            self.status.set("作成できません: " + str(e))
            return
        self.status.set("ファイルを作成しました: " + name)
        for pp in self.panes:
            if pp.scope_dir == p.scope_dir:
                pp.sel_target = path
                self.reload_pane(pp.idx)

    def reveal_in_explorer(self, iid):
        """文書内検索のヒット行 -> このアプリのエクスプローラでその場所を開く"""
        path = self.paths.get(iid, "")
        if not path:
            return
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            self.status.set("フォルダが存在しません: " + folder)
            return
        self.nb.select(self.tabf)
        self.sel_target = path
        self.navigate(folder)

    def on_tree_keypress(self, p, event):
        """一覧上で文字を打つと頭文字ジャンプ(Explorer準拠)。
        1秒以内に続けて打つと前方一致で絞り込み、同じ文字の連打は順送り"""
        ch = event.char
        if not ch or len(ch) != 1 or not ch.isprintable() or ch == " ":
            return None
        if event.state & 0x4 or event.state & 0x20000:   # Ctrl / Alt
            return None
        now = time.time()
        if now - self._ta_time > 1.0:
            self._ta_buf = ""
        self._ta_time = now
        t = p.tree
        kids = list(t.get_children())
        if not kids:
            return "break"
        cur = t.focus()
        start = kids.index(cur) if cur in kids else 0
        buf = self._ta_buf + ch.lower()
        same = (buf == buf[0] * len(buf))
        if same and len(buf) > 1:
            prefix = buf[0]
            order = kids[start + 1:] + kids[:start + 1]
        else:
            prefix = buf
            order = kids[start:] + kids[:start]
        for iid in order:
            if t.item(iid, "text").lower().startswith(prefix):
                t.selection_set(iid)
                t.focus(iid)
                t.see(iid)
                break
        self._ta_buf = buf
        return "break"

    def on_editor_change(self, event=None):
        self.save_conf()
        self.update_editor_note()

    def update_editor_note(self):
        exe = find_editor(self.var_editor.get())
        if exe:
            self.var_editor_note.set("空欄なら自動検出。現在の使用先: " + exe)
        else:
            self.var_editor_note.set("サクラエディタが見つかりません。"
                                     "sakura.exe のパスを指定してください")

    def browse_editor(self):
        fn = filedialog.askopenfilename(
            title="テキストエディタの実行ファイル",
            filetypes=[("実行ファイル", "*.exe"), ("すべて", "*.*")])
        if fn:
            self.var_editor.set(fn.replace("/", "\\"))
            self.on_editor_change()

    def restore_geometry(self, conf):
        g = conf.get("geom", "")
        m = re.match(r"^(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", g)
        if m:
            w, h, x, y = (int(v) for v in m.groups())
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            if 300 <= w <= sw + 50 and 200 <= h <= sh + 50 \
                    and -50 <= x < sw - 100 and -50 <= y < sh - 100:
                self.root.geometry(g)
        if conf.get("zoomed", "0") == "1":
            try:
                self.root.state("zoomed")
            except tk.TclError:
                pass

    def on_close(self):
        try:
            self.save_tab_state()
            self.save_tabs()
        except Exception:
            pass
        self.save_conf()
        self.root.destroy()

    # ---------- 実体を開く(共通) ----------
    def open_file_path(self, path):
        now = time.time()
        if path and now - self.last_open > 1.0 and hasattr(os, "startfile"):
            self.last_open = now
            try:
                os.startfile(path)
            except OSError:
                pass

    def open_folder_sel(self, path):
        now = time.time()
        if path and now - self.last_open > 1.0 and os.name == "nt":
            self.last_open = now
            try:
                subprocess.Popen('explorer /select,"%s"' % path)
            except OSError:
                pass

    def open_folder_plain(self, path):
        now = time.time()
        if path and now - self.last_open > 1.0 and os.name == "nt":
            self.last_open = now
            try:
                subprocess.Popen('explorer "%s"' % path)
            except OSError:
                pass

    # ---------- 設定の保存/読込 ----------
    def load_conf(self):
        conf = {}
        try:
            with open(CONF_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line:
                        k, v = line.split("=", 1)
                        conf[k.strip()] = v.strip()
        except OSError:
            pass
        return conf

    def save_conf(self):
        # 最大化中は通常時の位置/サイズを残す(書き込みで空にする前に読む)
        old_geom = self.load_conf().get("geom", "")
        try:
            with open(CONF_FILE, "w", encoding="utf-8") as f:
                f.write("theme=" + self.var_theme.get() + "\n")
                f.write("font=" + self.var_font.get() + "\n")
                f.write("size=" + self.var_fsize.get() + "\n")
                f.write("panes=" + ("2" if self.two_pane else "1") + "\n")
                f.write("editor=" + self.var_editor.get().strip() + "\n")
                f.write("dpath=" + ("1" if self.var_dpath.get() else "0")
                        + "\n")
                f.write("fpath=" + ("1" if self.var_fpath.get() else "0")
                        + "\n")
                f.write("fmode=" + self.find_mode() + "\n")
                f.write("fileidx=" + ("1" if self.var_fileidx.get() else "0")
                        + "\n")
                f.write("showtmp=" + ("1" if self.var_showtmp.get() else "0")
                        + "\n")
                f.write("fold=" + ("1" if self.var_fold.get() else "0")
                        + "\n")
                for key, tr in getattr(self, "col_trees", {}).items():
                    f.write(key + "=" + col_widths(tr) + "\n")
                try:
                    zoomed = self.root.state() == "zoomed"
                except tk.TclError:
                    zoomed = False
                f.write("zoomed=" + ("1" if zoomed else "0") + "\n")
                if not zoomed:
                    f.write("geom=" + self.root.winfo_geometry() + "\n")
                else:
                    f.write("geom=" + old_geom + "\n")
        except OSError:
            pass

    def on_conf_change(self, event=None):
        th = self.var_theme.get()
        if th == WIN_STD and self._last_theme != WIN_STD:
            if "Yu Gothic UI" in self.families:
                self.var_font.set("Yu Gothic UI")
            self.var_fsize.set("9")
        self._last_theme = th
        self.apply_theme(th)
        self.apply_font()
        self.save_conf()

    # ---------- テーマ/フォント ----------
    def apply_theme(self, name):
        raw = THEMES.get(name, THEMES[WIN_STD])
        p = {}
        for k, v in raw.items():
            if isinstance(v, str) and os.name != "nt":
                v = SYS_FALL.get(v, v)
            p[k] = v
        self.pal = p
        st = self.style
        if p["dark"]:
            themes = ("clam",)
        else:
            themes = ("vista", "xpnative", "clam")
        for t in themes:
            try:
                st.theme_use(t)
                break
            except tk.TclError:
                continue
        st.configure(".", background=p["bg"], foreground=p["fg"],
                     bordercolor=p["panel"], lightcolor=p["bg"],
                     darkcolor=p["bg"], focuscolor=p["accent"])
        st.configure("TFrame", background=p["bg"])
        st.configure("TPanedwindow", background=p["bg"])
        st.configure("TLabel", background=p["bg"], foreground=p["fg"])
        st.configure("Sub.TLabel", foreground=p["sub"])
        st.configure("Status.TLabel", background=p["panel"],
                     foreground=p["status"])
        st.configure("PaneHdr.TLabel", background=p["panel"],
                     foreground=p["sub"])
        st.configure("PaneHdrOn.TLabel", background=p["panel"],
                     foreground=p["status"])
        if p["dark"]:
            st.configure("TButton", background=p["panel"],
                         foreground=p["fg"], bordercolor=p["accent_dk"],
                         padding=(10, 4))
            st.map("TButton",
                   background=[("active", p["accent"])],
                   foreground=[("active", p["bg"]),
                               ("disabled", p["sub"])])
            st.configure("Accent.TButton", background=p["accent"],
                         foreground=p["bg"], bordercolor=p["gold"])
            st.map("Accent.TButton",
                   background=[("active", p["accent_dk"]),
                               ("disabled", p["panel"])],
                   foreground=[("active", p["fg"]),
                               ("disabled", p["sub"])])
            st.configure("Tool.TButton", background=p["bg"],
                         foreground=p["sub"], bordercolor=p["bg"],
                         padding=(6, 4))
            st.map("Tool.TButton",
                   background=[("active", p["panel"])],
                   foreground=[("active", p["accent"])])
            st.configure("TabBtn.TButton", background=p["panel"],
                         foreground=p["sub"], bordercolor=p["panel"],
                         padding=(10, 3))
            st.map("TabBtn.TButton",
                   background=[("active", p["bg"])],
                   foreground=[("active", p["accent"])])
        else:
            st.configure("TabBtn.TButton", padding=(8, 2))
        st.configure("TEntry", fieldbackground=p["field"],
                     foreground=p["fg"], bordercolor=p["accent_dk"],
                     insertcolor=p["insert"],
                     lightcolor=p["field"], darkcolor=p["field"])
        st.configure("TCombobox", fieldbackground=p["field"],
                     background=p["panel"], foreground=p["fg"],
                     arrowcolor=p["accent"], bordercolor=p["accent_dk"])
        st.map("TCombobox",
               fieldbackground=[("readonly", p["field"])],
               foreground=[("readonly", p["fg"])])
        st.configure("TNotebook", background=p["bg"], bordercolor=p["bg"],
                     tabmargins=(8, 6, 8, 0))
        st.configure("TNotebook.Tab", background=p["panel"],
                     foreground=p["sub"], padding=(18, 7),
                     bordercolor=p["panel"])
        st.map("TNotebook.Tab",
               background=[("selected", p["bg"])],
               foreground=[("selected", p["tabsel"])])
        st.configure("Treeview", background=p["field"],
                     fieldbackground=p["field"], foreground=p["fg"],
                     bordercolor=p["panel"])
        st.map("Treeview",
               background=[("selected", p["sel"])],
               foreground=[("selected", p["selfg"])])
        st.configure("Treeview.Heading", background=p["panel"],
                     foreground=p["sub"] if p["dark"] else p["fg"],
                     bordercolor=p["panel"],
                     relief="flat", padding=(8, 6))
        st.map("Treeview.Heading",
               foreground=[("active", p["accent"])],
               background=[("active", p["panel"])])
        st.configure("Side.Treeview", background=p["panel"],
                     fieldbackground=p["panel"], foreground=p["fg"],
                     bordercolor=p["panel"])
        st.map("Side.Treeview",
               background=[("selected", p["sel"])],
               foreground=[("selected", p["selfg"])])
        st.configure("Vertical.TScrollbar", background=p["panel"],
                     troughcolor=p["bg"], bordercolor=p["bg"],
                     arrowcolor=p["accent"])
        st.configure("Horizontal.TScrollbar", background=p["panel"],
                     troughcolor=p["bg"], bordercolor=p["bg"],
                     arrowcolor=p["accent"])
        self.root.configure(bg=p["bg"])
        self.root.option_add("*TCombobox*Listbox.background", p["field"])
        self.root.option_add("*TCombobox*Listbox.foreground", p["fg"])
        self.root.option_add("*TCombobox*Listbox.selectBackground",
                             p["sel"])
        self.root.option_add("*TCombobox*Listbox.selectForeground",
                             p["selfg"])
        for sp in self.spins:
            sp.configure(bg=p["field"], fg=p["fg"],
                         relief="flat" if p["dark"] else "sunken",
                         insertbackground=p["insert"],
                         buttonbackground=p["panel"],
                         readonlybackground=p["field"],
                         highlightthickness=0)
        for txt in self.texts:
            txt.configure(bg=p["field"], fg=p["fg"],
                          relief="flat" if p["dark"] else "sunken",
                          insertbackground=p["insert"],
                          highlightthickness=1 if p["dark"] else 0,
                          highlightbackground=p["gold"],
                          highlightcolor=p["accent"])
        for mn in self.all_menus():
            mn.configure(bg=p["panel"], fg=p["fg"],
                         activebackground=p["sel"],
                         activeforeground=p["selfg"])
        st.configure("TCheckbutton", background=p["bg"],
                     foreground=p["fg"], focuscolor=p["accent"])
        st.map("TCheckbutton",
               background=[("active", p["bg"])],
               foreground=[("active", p["accent"])])
        if p["dark"]:
            st.configure("TCheckbutton", indicatorbackground=p["field"],
                         indicatorforeground=p["accent"])
            st.map("TCheckbutton",
                   indicatorbackground=[("selected", p["field"]),
                                        ("active", p["panel"])])
        for line in self.gold_lines:
            line.configure(bg=p["gold"])
        for tr in (self.tree, self.panes[0].tree, self.panes[1].tree,
                   self.dtree):
            tr.tag_configure("even", background=p["field"])
            tr.tag_configure("odd", background=p["alt"])
            tr.tag_configure("err", foreground=p["gold"] if p["dark"]
                             else "#b06000")
            tr.tag_configure("folder", foreground=p["folder"])
        self.side.tag_configure("hdr", foreground=p["sub"])
        self.side.tag_configure("grp", foreground=p["sub"])
        self.side.tag_configure("itm", foreground=p["fg"])
        dark_title_bar(self.root, p["dark"])

    def all_menus(self):
        return (self.menu, self.fmenu, self.fmenu_term, self.tmenu,
                self.dmenu, self.dmenu_term)

    def apply_font(self):
        fam = self.var_font.get()
        try:
            size = int(self.var_fsize.get())
        except ValueError:
            size = 9
        size = max(8, min(20, size))
        self.var_fsize.set(str(size))
        f = (fam, size)
        self.cur_font = f
        st = self.style
        st.configure(".", font=f)
        st.configure("TButton", font=f)
        st.configure("Tool.TButton", font=(fam, size + 1))
        st.configure("TabBtn.TButton", font=f)
        st.configure("TNotebook.Tab", font=f)
        st.configure("Treeview", font=f)
        st.configure("Side.Treeview", font=f)
        st.configure("Treeview.Heading", font=f)
        st.configure("PaneHdr.TLabel", font=(fam, max(8, size - 1)))
        st.configure("PaneHdrOn.TLabel", font=(fam, max(8, size - 1),
                                               "bold"))
        lh = tkfont.Font(family=fam, size=size).metrics("linespace")
        st.configure("Treeview", rowheight=max(lh + 8, 20))
        st.configure("Side.Treeview", rowheight=max(lh + 6, 20))
        self.root.option_add("*TCombobox*Listbox.font",
                             "{%s} %d" % (fam, size))
        for wgt in self.font_widgets:
            wgt.configure(font=f)
        for sp in self.spins:
            sp.configure(font=f)
        for txt in self.texts:
            txt.configure(font=f)
        for mn in self.all_menus():
            mn.configure(font=f)
        st.configure("TCheckbutton", font=f)
        self.side.tag_configure("hdr", font=(fam, max(8, size - 1), "bold"))
        self.side.tag_configure("grp", font=(fam, max(8, size - 1), "bold"))
        self.side.tag_configure("itm", font=f)

    # ---------- お気に入り ----------
    def refresh_fav_text(self):
        self.fav_text.delete("1.0", "end")
        self.fav_text.insert(
            "1.0", "\n".join(k + "=" + v for k, v in self.favs))

    def save_fav_edit(self):
        favs = []
        for line in self.fav_text.get("1.0", "end").splitlines():
            line = line.strip()
            if "=" in line:
                k, v = line.split("=", 1)
                if k.strip() and v.strip():
                    favs.append((k.strip(), v.strip()))
        self.favs = favs
        save_favs(favs)
        self.refresh_fav_text()
        self.build_side()
        self.status.set("お気に入りを保存しました (%d件)" % len(favs))

    def show_fav_menu(self):
        self.nb.select(self.tab1)
        p = self.pal
        m = tk.Menu(self.root, tearoff=0, bg=p["panel"], fg=p["fg"],
                    activebackground=p["sel"], activeforeground=p["selfg"],
                    font=self.cur_font)
        if len(self.favs) > 100:
            m.add_command(label="件数が多い場合は左ペインの絞り込み"
                                " (Ctrl+B) が便利です",
                          state="disabled")
            m.add_separator()
        if self.favs:
            subs = {}
            for name, path in self.favs:
                g, label = fav_group(name)
                disp = path if len(path) <= 45 else path[:42] + "..."
                if g:
                    if g not in subs:
                        sm = tk.Menu(m, tearoff=0, bg=p["panel"],
                                     fg=p["fg"],
                                     activebackground=p["sel"],
                                     activeforeground=p["selfg"],
                                     font=self.cur_font)
                        subs[g] = sm
                        m.add_cascade(label=g, menu=sm)
                    subs[g].add_command(
                        label=label + "  —  " + disp,
                        command=lambda pth=path: self.use_favorite(pth))
                else:
                    m.add_command(
                        label=label + "  —  " + disp,
                        command=lambda pth=path: self.use_favorite(pth))
        else:
            m.add_command(label="（お気に入り未登録: F8で追加できます）",
                          state="disabled")
        m.add_separator()
        m.add_command(label="一覧の編集は「設定」タブで", state="disabled")
        x = self.btn_fav.winfo_rootx()
        y = self.btn_fav.winfo_rooty() + self.btn_fav.winfo_height()
        try:
            m.tk_popup(x, y)
        finally:
            m.grab_release()

    def use_favorite(self, path):
        self.var_dir.set(path)
        self.status.set("フォルダをセット: " + path)
        self.focus_widget(self.ent_word)

    def _add_fav(self, path):
        name = self.ask_fav(os.path.basename(path.rstrip("\\/")) or path)
        if not name:
            return
        self.favs.append((name, path))
        save_favs(self.favs)
        self.refresh_fav_text()
        self.build_side()
        self.status.set("お気に入りに追加: " + name)

    def add_favorite(self):
        if self.cur_tab() == TAB_FIND:
            path = self.find_focus_path()
            if path:
                self._add_fav(path)
            else:
                self.status.set("お気に入りに追加する行を選んでください")
            return
        if self.cur_tab() == TAB_EXP:
            path = None
            try:
                cur = self.root.focus_get()
            except (KeyError, tk.TclError):
                cur = None
            if cur in self.pane_by_tree:
                path = self.fpaths.get(self.ftree.focus())
            if not path:
                path = self.scope_dir
            if not path:
                self.status.set("お気に入りに追加する行を選ぶか、"
                                "フォルダに入ってから実行してください")
                return
            self._add_fav(path)
            return
        folder = self.var_dir.get().strip().strip('"')
        if not os.path.isdir(folder):
            self.status.set("フォルダ欄に有効なパスを入れてから追加してください")
            self.focus_widget(self.ent_dir)
            return
        self._add_fav(folder)

    def ask_fav(self, default_label):
        """お気に入り追加ダイアログ: グループ(既存選択/新規入力)+名前"""
        groups = []
        for n, _p in self.favs:
            g = fav_group(n)[0]
            if g and g not in groups:
                groups.append(g)
        dlg = tk.Toplevel(self.root)
        dlg.title("お気に入りに追加")
        dlg.configure(bg=self.pal["bg"])
        dlg.transient(self.root)
        dlg.geometry("+%d+%d" % (self.root.winfo_rootx() + 320,
                                 self.root.winfo_rooty() + 180))
        try:
            dlg.wait_visibility()
            dlg.grab_set()
        except tk.TclError:
            pass
        frm = ttk.Frame(dlg, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="グループ（▼から選択・新規は入力・空欄で直下）:"
                  ).grid(row=0, column=0, sticky="w")
        var_g = tk.StringVar(value=self.last_group)
        cb = ttk.Combobox(frm, textvariable=var_g, width=28,
                          values=groups, font=self.cur_font)
        cb.grid(row=1, column=0, sticky="we", pady=(2, 10))
        ttk.Label(frm, text="名前:").grid(row=2, column=0, sticky="w")
        var_n = tk.StringVar(value=default_label)
        ent = ttk.Entry(frm, width=30, textvariable=var_n,
                        font=self.cur_font)
        ent.grid(row=3, column=0, sticky="we", pady=(2, 0))
        res = {"v": None}

        def ok(event=None):
            label = var_n.get().strip().replace("=", " ")
            if not label:
                ent.focus_set()
                return
            g = var_g.get().strip().replace("=", " ")
            for sep in ("/", "／"):
                g = g.replace(sep, "-")
            self.last_group = g
            res["v"] = (g + "/" + label) if g else label
            dlg.destroy()

        def cancel(event=None):
            dlg.destroy()

        row = ttk.Frame(frm)
        row.grid(row=4, column=0, pady=(14, 0))
        ttk.Button(row, text="OK", command=ok).pack(side="left", padx=4)
        ttk.Button(row, text="キャンセル", command=cancel).pack(
            side="left", padx=4)
        cb.bind("<Return>", ok)
        ent.bind("<Return>", ok)
        dlg.bind("<Escape>", cancel)
        cb.focus_set()
        dlg.wait_window()
        return res["v"]

    def ask_text(self, title, prompt, default):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg=self.pal["bg"])
        dlg.transient(self.root)
        dlg.geometry("+%d+%d" % (self.root.winfo_rootx() + 340,
                                 self.root.winfo_rooty() + 200))
        try:
            dlg.wait_visibility()
            dlg.grab_set()
        except tk.TclError:
            pass
        ttk.Label(dlg, text=prompt).pack(padx=16, pady=(16, 4))
        var = tk.StringVar(value=default)
        ent = ttk.Entry(dlg, width=32, textvariable=var, font=self.cur_font)
        ent.pack(padx=16)
        res = {"v": None}

        def ok(event=None):
            res["v"] = var.get().strip()
            dlg.destroy()

        def cancel(event=None):
            dlg.destroy()

        row = ttk.Frame(dlg)
        row.pack(pady=12)
        ttk.Button(row, text="OK", command=ok).pack(side="left", padx=4)
        ttk.Button(row, text="キャンセル", command=cancel).pack(
            side="left", padx=4)
        ent.bind("<Return>", ok)
        dlg.bind("<Escape>", cancel)
        ent.focus_set()
        ent.select_range(0, "end")
        dlg.wait_window()
        return res["v"]

    def stop_scan(self):
        self.scan_stop.set()
        self.status.set("スキャン停止を要求しました")

    def show_help(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("ショートカット一覧 (F1)")
        dlg.configure(bg=self.pal["bg"])
        dlg.transient(self.root)
        dlg.geometry("+%d+%d" % (self.root.winfo_rootx() + 200,
                                 self.root.winfo_rooty() + 20))
        frm = tk.Frame(dlg, bg=self.pal["bg"])
        frm.pack(fill="both", expand=True, padx=8, pady=8)
        txt = tk.Text(frm, width=66, height=40, bg=self.pal["field"],
                      fg=self.pal["fg"], relief="flat",
                      font=self.cur_font, padx=12, pady=10)
        sb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.insert("1.0", HELP_TEXT)
        txt.configure(state="disabled")
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        dlg.bind("<Return>", lambda e: dlg.destroy())
        dlg.bind("<F1>", lambda e: dlg.destroy())
        dlg.focus_set()

    # ---------- フォーカス移動 ----------
    def focus_widget(self, wgt):
        self.nb.select(self.tab1)
        wgt.focus_set()
        try:
            wgt.select_range(0, "end")
            wgt.icursor("end")
        except (AttributeError, tk.TclError):
            pass

    def focus_tree(self):
        self.nb.select(self.tab1)
        kids = self.tree.get_children()
        if kids:
            cur = self.tree.focus() or kids[0]
            self.tree.selection_set(cur)
            self.tree.focus(cur)
            self.tree.see(cur)
        self.tree.focus_set()

    # ---------- 設定 ----------
    def get_exts(self):
        exts = []
        for it in self.ext_text.get("1.0", "end").split():
            it = it.strip().lower()
            if not it:
                continue
            if not it.startswith("."):
                it = "." + it
            exts.append(it)
        return tuple(exts)

    def save_ext(self):
        try:
            with open(EXT_FILE, "w", encoding="utf-8") as f:
                f.write(self.ext_text.get("1.0", "end").strip() + "\n")
            self.status.set("拡張子設定を保存しました: " + EXT_FILE)
        except OSError as e:
            self.status.set("保存に失敗: " + str(e))

    # ---------- 文書内検索 ----------
    def browse(self):
        d = filedialog.askdirectory()
        if d:
            self.var_dir.set(d)

    def start(self):
        if self.running:
            return
        self.nb.select(self.tab1)
        folder = self.var_dir.get().strip().strip('"')
        word = self.var_word.get()
        if not os.path.isdir(folder):
            self.status.set("フォルダが正しくありません")
            return
        if not word:
            self.status.set("検索文字列を入力してください")
            return
        try:
            ctx = int(self.var_ctx.get())
        except ValueError:
            ctx = CTX
        ctx = max(1, min(500, ctx))
        self.var_ctx.set(str(ctx))
        excludes = [s.strip().lower()
                    for s in self.var_ex.get().split(",") if s.strip()]
        exts = self.get_exts()
        kids = self.tree.get_children()
        if kids:
            self.tree.delete(*kids)
        self.paths = {}
        self.hit_count = 0
        self.row_i = 0
        self.running = True
        # 検索ごとに専用のキューと停止フラグを使う(停止直後に次の検索を
        # 始めても、前の検索の結果が混ざらないように)
        self.q = queue.Queue()
        self.doc_stop = threading.Event()
        self.btn.config(text="停止(Esc)")
        self.status.set("検索中...（Escで停止）")
        threading.Thread(target=search_worker,
                         args=(folder, word, ctx, excludes, exts, self.q,
                               self.doc_stop, bool(self.var_fold.get())),
                         daemon=True).start()
        self.root.after(100, self.poll, self.q)

    def toggle_doc_search(self):
        """検索ボタン: 検索中なら停止、そうでなければ検索開始"""
        if self.running:
            self.stop_doc_search()
        else:
            self.start()

    def stop_doc_search(self):
        if not self.running:
            return
        self.doc_stop.set()
        self.running = False
        self.btn.config(text="検索(F5)")
        self.status.set("検索を停止しました: ヒット %d 件（途中まで）"
                        % self.hit_count)

    def on_doc_escape(self, event=None):
        if self.running:
            self.stop_doc_search()
            return "break"
        return None

    def poll(self, q):
        if q is not self.q or not self.running:
            return                  # 停止済み/次の検索が始まった古い検索
        deadline = time.time() + 0.08
        wait = 100
        try:
            while True:
                if time.time() > deadline:
                    wait = 10       # まだ溜まっているのですぐ続きを処理
                    break
                item = q.get_nowait()
                if item[0] in ("hit", "err"):
                    _, path, name, folder, label, frag = item
                    if item[0] == "err":
                        tags = ("err",)
                    else:
                        tags = ("odd",) if self.row_i % 2 else ("even",)
                    self.row_i += 1
                    ikey = file_type(name)[0]
                    iid = self.tree.insert(
                        "", "end", text=name, image=self.icons[ikey],
                        values=(folder, label, frag), tags=tags)
                    self.paths[iid] = path
                    if item[0] == "hit":
                        self.hit_count += 1
                        self.status.set("検索中... ヒット %d 件（Escで停止）"
                                        % self.hit_count)
                else:
                    _, nf, nh, ne = item
                    if not self.running:
                        return      # 停止ボタンで打ち切り済み
                    self.status.set(
                        "完了: 対象 %d ファイル / ヒット %d 件 / エラー %d 件"
                        % (nf, nh, ne))
                    self.running = False
                    self.btn.config(text="検索(F5)")
                    if nh > 0:
                        self.focus_tree()
                    return
        except queue.Empty:
            pass
        if self.running:
            self.root.after(wait, self.poll, q)

    def export_csv(self):
        rows = self.tree.get_children()
        if not rows:
            self.status.set("出力する結果がありません")
            return
        fn = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="search_result.csv",
            filetypes=[("CSVファイル", "*.csv")])
        if not fn:
            return
        with open(fn, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["ファイル名", "フォルダ", "場所", "前後の文脈", "フルパス"])
            for iid in rows:
                v = self.tree.item(iid, "values")
                name = self.tree.item(iid, "text")
                w.writerow([name] + list(v) + [self.paths.get(iid, "")])
        self.status.set("CSVを出力しました: " + fn)

    def open_file_of(self, iid):
        self.open_file_path(self.paths.get(iid))

    def open_folder_of(self, iid):
        self.open_folder_sel(self.paths.get(iid))

    def on_dblclick(self, event):
        """ダブルクリックでファイルを開く(以前は1クリックで開いてしまい、
        行を選ぶだけのつもりでもファイルが起動していた)。
        列見出しの境界なら列幅を内容に合わせる"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            autofit_column(self.tree, self.tree.identify_column(event.x),
                           self.cur_font)
        elif region in ("tree", "cell"):
            self.open_file_of(self.tree.identify_row(event.y))
        return "break"

    def on_tree_folder(self, event):
        self.open_folder_of(self.tree.focus())
        return "break"

    def on_tree_file(self, event):
        self.open_file_of(self.tree.focus())
        return "break"

    def on_rclick(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.menu_iid = iid
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def on_menu_key(self, event):
        iid = self.tree.focus()
        if not iid:
            return "break"
        self.menu_iid = iid
        bbox = self.tree.bbox(iid)
        if bbox:
            x = self.tree.winfo_rootx() + bbox[0] + 60
            y = self.tree.winfo_rooty() + bbox[1] + bbox[3]
        else:
            x = self.tree.winfo_rootx() + 60
            y = self.tree.winfo_rooty() + 20
        try:
            self.menu.tk_popup(x, y)
        finally:
            self.menu.grab_release()
        return "break"

    def copy_path(self, kind):
        path = self.paths.get(self.menu_iid, "")
        if not path:
            return
        if kind == "dir":
            path = os.path.dirname(path)
        elif kind == "name":
            path = os.path.basename(path)
        self.to_clip(path)


def main():
    set_app_id()
    root = tk.Tk()
    app = App(root)

    def show_err(exc, val, tb):
        try:
            app.status.set("エラー: " + str(val))
        except Exception:
            pass
    root.report_callback_exception = show_err
    root.mainloop()


if __name__ == "__main__":
    if "--scan" in sys.argv:
        cli_scan()
    elif "--make-icon" in sys.argv:
        ico_path = os.path.join(BASE_DIR, "doc_search.ico")
        write_ico(ico_path)
        print("アイコンを書き出しました: " + ico_path)
        print("exe化の例: pyinstaller --onefile --noconsole "
              "--icon=doc_search.ico doc_search.py")
    else:
        main()
