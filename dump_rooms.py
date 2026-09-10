"""倒出「檢查地點及檢查室設定」目前選取地點底下的檢查室（唯讀）。

這個畫面是主從結構：
    表 M = 檢查地點（31 個，含院區別）
    表 D = 只顯示「目前選取的那個地點」底下的檢查室
所以一次只倒得到一支，要全部就得在 RIS 逐一點過。

檔名會自動帶上地點代碼，各地點不會互相覆蓋 ——
目前選哪一列是靠「哪一格有鍵盤焦點」判斷的，
判斷不出來就不亂猜，直接停下來請人確認（標錯的資料比沒有資料更危險）。

*** 這支程式不會修改任何東西，也不會點畫面。***

用法：在 RIS 選好一個檢查地點，然後
    python dump_rooms.py
"""

import csv
import os
import re
import sys

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
MASTER_ID = "M"
DETAIL_ID = "D"
OUT_DIR = "檢查室"


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def connect():
    hwnd = None
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                hwnd = w.handle
                break
        except Exception:
            continue
    if not hwnd:
        sys.exit("找不到 RIS 主視窗。")
    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    return app.window(handle=hwnd)


def tables(win):
    out = {}
    for t in win.descendants(control_type="Table"):
        out[str(t.element_info.automation_id)] = t
    return out


def read_grid(grid):
    """(欄位, 每列的值)。欄名從格子名稱推，空的那格是列首要跳過。"""
    rows = []
    cols, keep = [], None
    data = []
    for c in grid.children():
        name = c.window_text()
        if name.startswith("資料列"):
            try:
                data.append((int(name.split()[1]), c))
            except (IndexError, ValueError):
                continue
    data.sort()
    for idx, r in data:
        kids = r.children()
        if not kids:
            continue
        if keep is None:
            names = [re.sub(r"\s*資料列\s*\d+$", "", k.window_text()).strip()
                     for k in kids]
            keep = [i for i, n in enumerate(names) if n]
            cols = [names[i] for i in keep]
        vals = []
        for i in keep:
            if i >= len(kids):
                vals.append("")
                continue
            k = kids[i]
            v = None
            try:
                v = k.legacy_properties().get("Value")
            except Exception:
                pass
            if v in (None, ""):
                try:
                    v = k.iface_value.CurrentValue
                except Exception:
                    v = k.window_text()
            vals.append(norm(v))
        if all(x in ("", "(null)") for x in vals):
            continue
        rows.append((idx, vals))
    return cols, rows


def selected_row(grid):
    """哪一列是目前選取的？靠「哪一格有鍵盤焦點」判斷。找不到回 None。"""
    for r in grid.children():
        if not r.window_text().startswith("資料列"):
            continue
        for k in r.children():
            try:
                if k.element_info.element.CurrentHasKeyboardFocus:
                    return r
            except Exception:
                continue
    return None


def main():
    win = connect()
    tabs = tables(win)
    if MASTER_ID not in tabs or DETAIL_ID not in tabs:
        sys.exit("這個畫面沒有主從兩張表，請確認開的是「檢查地點及檢查室設定」。")

    mcols, mrows = read_grid(tabs[MASTER_ID])
    cur = selected_row(tabs[MASTER_ID])
    if cur is None:
        sys.exit("判斷不出目前選的是哪個地點。請在主表的資料格上點一下再跑一次。")

    idx = int(cur.window_text().split()[1])
    row = dict(zip(mcols, next((v for i, v in mrows if i == idx), [])))
    code = row.get("檢查地點代碼", "")
    name = row.get("檢查地點名稱", "")
    camp = row.get("院區別", "")
    if not code:
        sys.exit("讀不到地點代碼，停下來以免存成不知道是誰的檔案。")

    dcols, drows = read_grid(tabs[DETAIL_ID])
    os.makedirs(OUT_DIR, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "_", f"{code}_{name}")
    out = os.path.join(OUT_DIR, f"檢查室_{safe}.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["檢查地點代碼", "檢查地點名稱", "院區別"] + dcols)
        for _i, vals in drows:
            w.writerow([code, name, camp] + vals)

    print(f"地點：{code}  {name}  ({camp})")
    print(f"檢查室 {len(drows)} 間 -> {out}")
    for _i, vals in drows[:6]:
        print("   " + " | ".join(vals[:3]))
    if len(drows) > 6:
        print(f"   ...（共 {len(drows)} 間）")

    done = len(os.listdir(OUT_DIR))
    print(f"\n已完成 {done} / {len(mrows)} 個地點")


if __name__ == "__main__":
    main()
