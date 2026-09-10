"""
RIS 醫師班表維護 - 讀取表格實際內容
====================================
純讀取。不點擊、不輸入、不存檔、不修改任何資料。
唯一的動作是把畫面上已經查出來的表格內容抄下來。

用法：
  1. 開著「醫師班表維護」，並已查詢出資料
  2. python 3_read_table.py
  3. 產生 table.tsv（可用 Excel 開）和 table_debug.txt，兩個都給我看
"""

import io
import re
import sys

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"
OUT_TSV = "table.tsv"
OUT_DBG = "table_debug.txt"

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def cell_value(cell):
    """一格的實際內容可能藏在好幾個地方，逐一嘗試。回傳 (值, 來源)。"""
    # 1) ValuePattern —— 可編輯格子最標準的來源
    try:
        v = cell.iface_value.CurrentValue
        if v not in (None, ""):
            return v, "value"
    except Exception:
        pass
    # 2) LegacyIAccessible.Value —— WinForms 表格最常見的來源
    try:
        props = cell.legacy_properties()
        v = props.get("Value")
        if v not in (None, ""):
            return v, "legacy"
        v = props.get("Description")
        if v not in (None, ""):
            return v, "legacy-desc"
    except Exception:
        pass
    # 3) 底下如果還有子元件，值可能在子元件上
    try:
        kids = cell.children()
        if kids:
            texts = [k.window_text() for k in kids if k.window_text()]
            if texts:
                return " ".join(texts), "child"
    except Exception:
        pass
    # 4) 最後才退回名牌
    try:
        return cell.window_text(), "name"
    except Exception:
        return "", "none"


def main():
    hwnd = None
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                hwnd = w.handle
                out(f"主視窗: {w.window_text()}")
                break
        except Exception:
            continue
    if not hwnd:
        out("找不到 RIS 主視窗，請確認程式開著。")
        return

    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    main_win = app.window(handle=hwnd)
    grid = main_win.child_window(auto_id=FORM_AUTO_ID).child_window(auto_id=GRID_AUTO_ID)
    grid.wait("exists", timeout=15)

    rows = grid.children()
    out(f"表格共 {len(rows)} 個元件\n")

    # 找出標題列和資料列
    header = None
    data_rows = []
    for r in rows:
        try:
            name = r.window_text()
        except Exception:
            continue
        if "標題" in name or "上方資料列" in name:
            header = r
        elif name.startswith("資料列"):
            data_rows.append(r)

    # 標題
    cols = []
    if header is not None:
        for c in header.children():
            t = c.window_text()
            if "左上方" not in t:
                cols.append(t)
    out(f"欄位（{len(cols)}）: {cols}\n")

    # 前 2 列做詳細診斷：每一格試過的來源都印出來
    out("=" * 60)
    out("診斷：前 2 列每一格的讀取來源")
    out("=" * 60)
    for r in data_rows[:2]:
        out(f"\n--- {r.window_text()} ---")
        for c in r.children():
            val, src = cell_value(c)
            out(f"  [{src:11}] {c.window_text():28} = {val!r}")

    # 全部抄下來
    out("\n" + "=" * 60)
    out(f"開始讀取全部 {len(data_rows)} 列")
    out("=" * 60)

    table = [cols]
    for i, r in enumerate(data_rows):
        vals = []
        for c in r.children():
            if "資料列" == c.window_text().split()[0] and len(c.window_text().split()) == 2:
                continue  # 跳過列首的序號格
            v, _ = cell_value(c)
            vals.append(v)
        table.append(vals)
        if (i + 1) % 10 == 0:
            print(f"  ...已讀 {i + 1}/{len(data_rows)} 列")

    with open(OUT_TSV, "w", encoding="utf-8-sig") as f:
        for row in table:
            f.write("\t".join(str(x).replace("\t", " ") for x in row) + "\n")

    out(f"\n完成，共 {len(table) - 1} 列資料 -> {OUT_TSV}")
    out("\n--- 前 5 列預覽 ---")
    for row in table[:6]:
        out("  " + " | ".join(str(x) for x in row))

    with open(OUT_DBG, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f">>> 診斷檔 {OUT_DBG}")


if __name__ == "__main__":
    main()
