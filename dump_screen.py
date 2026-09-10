"""
把 RIS 目前開著的設定畫面整份倒成 CSV（唯讀）
==============================================
RIS 每個設定畫面（工作時段設定、使用者帳號設定、醫師群組分派檔…）
用的都是同一種表格元件：

    Pane aid=f???? name='畫面名稱'
      └─ Table aid=D name='DataGridView'

所以不必每個畫面各寫一支，認 auto_id 'D' 就行。
欄位也不寫死，從第一列的格子名稱自己推出來。

*** 這支程式不會修改任何東西。***
只讀元件屬性，不點擊、不輸入、不存檔。

用法：
    在 RIS 切到想倒的那個設定畫面，然後
        python dump_screen.py
    會存成 傾印_<畫面名稱>.csv

    python dump_screen.py --list     # 只列出目前有哪些畫面可以倒
"""

import argparse
import csv
import re
import sys
from datetime import datetime

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
GRID_AUTO_ID = "D"


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
        sys.exit("找不到 RIS 主視窗，請確認 RIS 開著。")
    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    return app.window(handle=hwnd)


def _ancestor_form(el):
    """從一個元件往上找它屬於哪個設定畫面。回傳 (畫面名稱, 畫面元件)。

    不同畫面的表格藏的深度不一樣（工作時段設定是直接子元件，
    使用者設定檔則包在 SplitContainer 裡好幾層），所以不能寫死路徑，
    改成從表格往上找「auto_id 以 f 開頭又有名字」的那一層。
    """
    for _ in range(8):
        try:
            el = el.parent()
        except Exception:
            return "未知畫面", None
        if el is None:
            return "未知畫面", None
        info = el.element_info
        if str(info.automation_id or "").startswith("f") and norm(info.name):
            return norm(info.name), el
    return "未知畫面", None


def find_tables(win):
    """找出畫面上所有表格。回傳 [(畫面名稱, 表格 auto_id, 表格元件, 畫面元件), ...]

    RIS 一次只把作用中的畫面放進元件樹，所以通常只會找到目前開著的那個。
    一個畫面可能有不只一張表（例如使用者設定檔有帳號清單和角色明細兩張）。
    """
    out = []
    for t in win.descendants(control_type="Table"):
        name, form = _ancestor_form(t)
        out.append((name, str(t.element_info.automation_id), t, form))
    return out


def campus_of(form):
    """讀出「查詢條件」裡的院區。回傳 "" 表示這個畫面沒有院區條件。

    標錯院區的資料比沒有資料更危險，所以判斷不出來就回空字串，不猜。

    各畫面的院區欄位不一樣：
      使用者設定檔  fUser           兩個：Q_HOSPITALCODE(查詢區) + HOSPITALCODE(明細區)
      醫師週班設定  fWeekWorkSheet  一個：HOSPITALCODE(它就是查詢條件)

    明細區那個顯示的是「目前選到的那一筆」的院區，換院區查詢後不會馬上更新；
    取到它，檔名就會標成上一次的院區 —— 實際踩過（2026-07-23 查金山存成竹東）。

    所以規則是：
      有 Q_ 前綴  -> 那一定是查詢條件，用它
      沒有 Q_ 但整個畫面只有一個 -> 沒得混淆，用它
      沒有 Q_ 卻有好幾個         -> 分不出來，回空字串
    """
    if form is None:
        return ""
    try:
        combos = list(form.descendants(control_type="ComboBox"))
    except Exception:
        return ""

    def value(c):
        try:
            return norm(c.iface_value.CurrentValue)
        except Exception:
            return norm(c.window_text())

    hosp = [c for c in combos
            if "HOSPITAL" in str(c.element_info.automation_id or "").upper()]
    if not hosp:
        return ""
    q = [c for c in hosp
         if str(c.element_info.automation_id or "").upper().startswith("Q_")]
    if q:
        return value(q[0])
    if len(hosp) == 1:
        return value(hosp[0])
    return ""


def read_grid(grid):
    """把表格讀成 (欄位名稱, 資料列)。欄位從第一列的格子名稱推出來。

    格子的名稱長這樣：'排班時段 資料列 0'，去掉尾巴的「資料列 N」就是欄名。
    有的表格第一格是列首（名稱只有 '資料列 0'，去掉之後是空的），
    有的表格沒有列首（第一格就是資料）—— 所以用「去掉之後空不空」來判斷，
    不能寫死跳過第一格，否則會把第一欄吃掉。
    """
    data_rows = []
    for c in grid.children():
        name = c.window_text()
        if name.startswith("資料列"):
            try:
                data_rows.append((int(name.split()[1]), c))
            except (IndexError, ValueError):
                continue
    data_rows.sort()

    cols, rows, keep = [], [], None
    for idx, r in data_rows:
        kids = r.children()
        if not kids:
            continue
        if keep is None:
            names = [re.sub(r"\s*資料列\s*\d+$", "", k.window_text()).strip()
                     for k in kids]
            keep = [i for i, n in enumerate(names) if n]   # 名字空的是列首
            cols = [names[i] for i in keep]
        kids = [kids[i] for i in keep if i < len(kids)]
        vals = []
        for k in kids:
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
            continue                       # 最底下那列空白的新增列
        rows.append(vals)
    return cols, rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true", help="只列出可以倒的畫面")
    args = ap.parse_args()

    win = connect()
    tables = find_tables(win)
    if not tables:
        sys.exit("畫面上找不到表格。請先在 RIS 切到要倒的那一個設定畫面。")

    print(f"找到 {len(tables)} 張表格：")
    for name, aid, t, form in tables:
        camp = campus_of(form)
        print(f"  - {name} / 表格 {aid}" + (f" / 院區 {camp}" if camp else ""))
    if args.list:
        return

    multi = len({n for n, _a, _t, _f in tables}) < len(tables)
    for name, aid, grid, form in tables:
        cols, rows = read_grid(grid)
        if not cols:
            print(f"\n[{name}/{aid}] 讀不到欄位，跳過。")
            continue
        camp = campus_of(form)
        out = ("傾印_" + name
               + (f"_{camp}" if camp else "")
               + (f"_{aid}" if multi else "") + ".csv")
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(cols)
            w.writerows(rows)
        print(f"\n[{name}] {len(rows)} 列 × {len(cols)} 欄 -> {out}")
        print("  欄位:", " / ".join(cols))
        for r in rows[:5]:
            print("   ", " | ".join(r))
        if len(rows) > 5:
            print(f"    ...（共 {len(rows)} 列）")

    print(f"\n完成 {datetime.now():%H:%M:%S}。這支程式沒有修改任何東西。")


if __name__ == "__main__":
    main()
