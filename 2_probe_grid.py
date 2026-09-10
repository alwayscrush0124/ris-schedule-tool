"""
RIS 醫師班表維護 - 表格讀取測試
=================================
目的：測試有沒有辦法讀到查詢結果表格裡的內容。

安全性說明：
  * 階段 A（自動執行）：純讀取。只讀畫面上的欄位值，不點、不填、不存檔。
  * 階段 B（要你打字確認才會做）：會對表格送出「全選 + 複製」。
      這兩個動作不會修改任何資料，但它確實會操作到畫面，所以我讓你自己決定。
  * 全程不會碰「存檔」。

用法：
  1. 開著 RIS 的「醫師班表維護」，並且已經查詢出一批資料
  2. python 2_probe_grid.py
  3. 把 grid_probe.txt 給我看
"""

import io
import re
import sys
import time

from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

OUT = "grid_probe.txt"
MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def find_main_hwnd():
    """用 win32 找到 RIS 主視窗的代碼（handle）。這是最可靠的一步。"""
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                return w.handle, w.window_text()
        except Exception:
            continue
    return None, None


def stage_a_read_fields(hwnd):
    """階段 A：純讀取查詢條件欄位。"""
    out("\n########## 階段 A：讀取查詢條件（純讀取） ##########")
    app = Application(backend="win32").connect(handle=hwnd)
    main = app.window(handle=hwnd)

    wanted = [
        ("院區別", "HOSPITALCODE"),
        ("醫師類別", "USERLEVELCODE"),
        ("排班醫師", "USERCODE"),
        ("起始日期", "STARTDATE"),
        ("結束日期", "ENDDATE"),
        ("排班類別", "SCHEDULETYPECODE1"),
        ("統計用分類", "STATISTICGROUPCODE"),
    ]
    for label, auto_id in wanted:
        try:
            ctl = main.child_window(auto_id=auto_id)
            out(f"  {label:12} ({auto_id:20}) = {ctl.window_text()!r}")
        except Exception as e:
            out(f"  {label:12} ({auto_id:20}) 讀取失敗: {type(e).__name__}")


def stage_a2_try_uia(hwnd):
    """階段 A2：改用 handle 連 UIA，測試能不能讀到表格內容。"""
    out("\n########## 階段 A2：嘗試用 UIA 讀表格（純讀取） ##########")
    try:
        app = Application(backend="uia").connect(handle=hwnd, timeout=20)
        main = app.window(handle=hwnd)
        out("  UIA 連線成功")
    except Exception as e:
        out(f"  !! UIA 連線失敗: {type(e).__name__}: {e}")
        out("  -> 這條路不通，看階段 B 的結果。")
        return

    try:
        form = main.child_window(auto_id=FORM_AUTO_ID)
        grid = form.child_window(auto_id=GRID_AUTO_ID)
        grid.wait("exists", timeout=15)
        out("  找到表格元件")
    except Exception as e:
        out(f"  !! 找不到表格: {type(e).__name__}: {e}")
        return

    # 試各種讀法，哪個成功都好
    for name, fn in [
        ("texts()", lambda: grid.texts()),
        ("item_count()", lambda: grid.item_count()),
        ("column_count()", lambda: grid.column_count()),
        ("children 數量", lambda: len(grid.children())),
    ]:
        try:
            val = fn()
            s = repr(val)
            out(f"  {name:16} -> {s[:400]}{' ...(截斷)' if len(s) > 400 else ''}")
        except Exception as e:
            out(f"  {name:16} -> 失敗 ({type(e).__name__})")

    # 逐格讀取前 3 列，確認格子讀得到
    try:
        out("\n  --- 嘗試逐格讀取前 3 列 ---")
        rows = grid.children()
        for r in rows[:4]:
            cells = [c.window_text() for c in r.children()]
            out(f"    {cells}")
    except Exception as e:
        out(f"  逐格讀取失敗: {type(e).__name__}: {e}")


def stage_b_clipboard(hwnd):
    """階段 B：對表格送 Ctrl+A / Ctrl+C，從剪貼簿讀回整張表。"""
    out("\n########## 階段 B：剪貼簿讀取 ##########")
    try:
        import pyperclip
    except ImportError:
        out("  需要 pyperclip： pip install pyperclip")
        return

    print("\n" + "=" * 60)
    print("階段 B 會做這兩件事：")
    print("  1. 把焦點移到查詢結果的表格上")
    print("  2. 送出 Ctrl+A（全選）和 Ctrl+C（複製）")
    print("這兩個動作不會修改資料，也不會存檔。")
    print("你目前的剪貼簿內容會被覆蓋掉。")
    print("=" * 60)
    ans = input("確定要執行嗎？請完整輸入 yes： ").strip().lower()
    if ans != "yes":
        out("  你選擇跳過階段 B。")
        return

    before = ""
    try:
        before = pyperclip.paste()
    except Exception:
        pass

    app = Application(backend="win32").connect(handle=hwnd)
    main = app.window(handle=hwnd)
    grid = main.child_window(auto_id=GRID_AUTO_ID)

    main.set_focus()
    time.sleep(0.4)
    try:
        grid.set_focus()
    except Exception:
        # 有些控制項不吃 set_focus，改成點左上角一格
        rect = grid.rectangle()
        grid.click_input(coords=(30, 30))
    time.sleep(0.4)

    send_keys("^a")
    time.sleep(0.3)
    send_keys("^c")
    time.sleep(0.6)

    try:
        data = pyperclip.paste()
    except Exception as e:
        out(f"  讀剪貼簿失敗: {e}")
        return

    if not data or data == before:
        out("  !! 剪貼簿沒有變化 —— 這個表格可能不支援複製。")
        return

    lines = data.splitlines()
    out(f"  成功！複製到 {len(lines)} 行")
    out("  --- 前 8 行（欄位以 | 分隔）---")
    for ln in lines[:8]:
        out("    " + " | ".join(ln.split("\t")))
    out(f"\n  --- 完整內容 ({len(data)} 字元) ---")
    out(data)


def main():
    hwnd, title = find_main_hwnd()
    if not hwnd:
        out("找不到 RIS 主視窗。請確認 RIS 開著，然後重跑。")
        return
    out(f"找到主視窗: {title}")
    out(f"handle = {hwnd}")

    stage_a_read_fields(hwnd)
    stage_a2_try_uia(hwnd)
    stage_b_clipboard(hwnd)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f"\n>>> 已存成 {OUT}")


if __name__ == "__main__":
    main()
