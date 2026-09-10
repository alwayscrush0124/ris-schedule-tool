"""
RIS 表格格子 - 編輯方式探測
============================
目的：搞清楚「排班醫師」那一格要怎麼正確地填值。
      上一版用「打字」的方式失敗了（格子要的是帳號，不是顯示的那串字），
      這支改成研究「點開下拉選單用選的」這條路。

會做什麼：
  1. 讀那一格目前的值，以及它支援哪些操作方式（純讀取）
  2. 詢問你之後，點兩下進入編輯模式，把下拉選單的選項列出來
  3. 按 ESC 取消編輯 —— 不會留下任何修改

*** 不會存檔。ESC 取消後畫面回到原狀。***

用法：
  python 8_probe_cell.py 0        # 0 是表格列號
"""

import argparse
import io
import re
import sys
import time

from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"
OUT = "cell_probe.txt"

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("row", type=int, help="表格列號（從 0 開始）")
    args = ap.parse_args()

    hwnd = None
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                hwnd = w.handle
                break
        except Exception:
            continue
    if not hwnd:
        out("找不到 RIS 主視窗。")
        return

    app_uia = Application(backend="uia").connect(handle=hwnd, timeout=20)
    main_win = app_uia.window(handle=hwnd)
    form = main_win.child_window(auto_id=FORM_AUTO_ID)
    grid = form.child_window(auto_id=GRID_AUTO_ID)
    grid.wait("exists", timeout=15)

    # 找到那一列的「排班醫師」格
    target_cell = None
    for r in grid.children():
        name = r.window_text()
        if name.startswith("資料列") and name.split()[1:] == [str(args.row)]:
            for c in r.children():
                if c.window_text().startswith("排班醫師"):
                    target_cell = c
            break
    if target_cell is None:
        out(f"找不到第 {args.row} 列的排班醫師格。")
        return

    out("=" * 70)
    out(f"第 {args.row} 列的「排班醫師」格")
    out("=" * 70)

    # --- 純讀取的部分 ---
    try:
        out(f"  目前的值 (ValuePattern) = {target_cell.iface_value.CurrentValue!r}")
        out(f"  是否唯讀                = {target_cell.iface_value.CurrentIsReadOnly}")
    except Exception as e:
        out(f"  ValuePattern 讀取失敗: {e}")

    try:
        lp = target_cell.legacy_properties()
        out(f"  LegacyIAccessible Value = {lp.get('Value')!r}")
        out(f"  LegacyIAccessible Name  = {lp.get('Name')!r}")
        out(f"  DefaultAction           = {lp.get('DefaultAction')!r}")
    except Exception as e:
        out(f"  legacy 讀取失敗: {e}")

    # 支援哪些操作方式
    out("\n  支援的操作方式：")
    for pat in ["value", "selectionitem", "expandcollapse", "invoke",
                "legacyiaccessible", "griditem", "toggle"]:
        try:
            getattr(target_cell, f"iface_{pat}")
            out(f"    O {pat}")
        except Exception:
            out(f"    -  {pat}")

    # --- 需要互動的部分 ---
    print("\n" + "=" * 60)
    print("接下來會：點兩下那一格 -> 讀出下拉選單的選項 -> 按 ESC 取消")
    print("ESC 會取消編輯，不會留下修改，也不會存檔。")
    print("=" * 60)
    if input("要繼續嗎？請完整輸入 yes： ").strip().lower() != "yes":
        out("\n你選擇跳過互動測試。")
        _save()
        return

    main_win.set_focus()
    time.sleep(0.3)
    try:
        target_cell.double_click_input()
    except Exception as e:
        out(f"\n點擊失敗: {e}")
        _save()
        return
    time.sleep(0.8)

    # 進入編輯模式後，會冒出一個編輯用的控制項
    out("\n  --- 進入編輯模式後新出現的控制項 ---")
    found_items = None
    try:
        for d in form.descendants():
            try:
                ct = str(d.element_info.control_type)
                cls = d.element_info.class_name or ""
            except Exception:
                continue
            if "Editing" in cls or "ComboBox" in cls or "Editing" in ct:
                out(f"    class={cls}  control_type={ct}  text={d.window_text()!r}")
                # 用 win32 直接問選單內容，不需要展開
                try:
                    app_w32 = Application(backend="win32").connect(handle=hwnd)
                    combo = app_w32.window(handle=d.element_info.handle)
                    items = combo.item_texts()
                    if items:
                        found_items = items
                except Exception:
                    pass
    except Exception as e:
        out(f"    列舉失敗: {e}")

    if found_items:
        out(f"\n  下拉選單有 {len(found_items)} 個選項，前 30 個：")
        for i, t in enumerate(found_items[:30]):
            out(f"    [{i:3}] {t!r}")
    else:
        out("\n  沒讀到下拉選單的選項。")

    # 取消編輯
    send_keys("{ESC}")
    time.sleep(0.4)
    send_keys("{ESC}")
    out("\n  已按 ESC 取消編輯。")

    try:
        out(f"  取消後這一格的值 = {target_cell.iface_value.CurrentValue!r}")
    except Exception:
        pass

    _save()


def _save():
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f"\n>>> 已存成 {OUT}")


if __name__ == "__main__":
    main()
