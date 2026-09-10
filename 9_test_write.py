"""
RIS 表格格子 - 寫入格式測試
============================
目的：確認「排班醫師」那一格到底接受什麼格式的值。
      顯示出來的是 'CCTAOFF均分'，但實際要填的可能只是帳號 'CCTAOFF'。

會做什麼：
  依序試你指定的候選值，每試一個就讀回來看格子變成什麼。
  一旦有一個成功就停下來。

*** 不會存檔。***
但要注意：設值是會立刻生效在畫面上的，ESC 救不回來。
如果全部試完都失敗，格子可能會變成空的 ——
用「重新查詢 -> 跳出提示選『否』」就能還原。

用法：
  python 9_test_write.py 0 --expect CCTAOFF CCTAOFF "CCTAOFF均分"
                          ^列號        ^期待看到的帳號  ^要試的候選值們
"""

import argparse
import io
import re
import time

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"
OUT = "write_test.txt"

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("row", type=int)
    ap.add_argument("--expect", required=True,
                    help="成功的話格子裡應該要出現的關鍵字，例如 CCTAOFF")
    ap.add_argument("candidates", nargs="+", help="要依序嘗試的值")
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

    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    grid = (app.window(handle=hwnd)
            .child_window(auto_id=FORM_AUTO_ID)
            .child_window(auto_id=GRID_AUTO_ID))
    grid.wait("exists", timeout=15)

    cell = None
    for r in grid.children():
        name = r.window_text()
        if name.startswith("資料列") and name.split()[1:] == [str(args.row)]:
            for c in r.children():
                if c.window_text().startswith("排班醫師"):
                    cell = c
            break
    if cell is None:
        out(f"找不到第 {args.row} 列的排班醫師格。")
        return

    original = cell.iface_value.CurrentValue
    out(f"第 {args.row} 列，目前的值 = {original!r}")
    out(f"要試的候選值: {args.candidates}")
    out(f"成功判斷標準: 格子裡出現 {args.expect!r}\n")

    print("=" * 60)
    print("這會實際改動畫面上那一格（但不存檔）。")
    print("試錯過程中格子可能暫時變成空白，屬正常。")
    print("結束後如果要還原：重新查詢 -> 提示選「否」。")
    print("=" * 60)
    if input("要開始嗎？請完整輸入 yes： ").strip().lower() != "yes":
        out("已取消。")
        return

    win = None
    for cand in args.candidates:
        out(f"\n--- 試 {cand!r} ---")
        try:
            cell.iface_value.SetValue(cand)
        except Exception as e:
            out(f"  設值時發生錯誤: {type(e).__name__}: {e}")
            continue
        time.sleep(0.4)
        try:
            now = cell.iface_value.CurrentValue
        except Exception as e:
            out(f"  讀回失敗: {e}")
            continue
        out(f"  格子現在是 {now!r}")
        if args.expect in str(now):
            out(f"\n  >>> 成功！正確的寫入格式是 {cand!r}")
            win = cand
            break
        else:
            out("  不對，繼續試下一個")

    out("\n" + "=" * 60)
    if win:
        out(f"結論：要寫入的值是 {win!r}，顯示出來會變成 {cell.iface_value.CurrentValue!r}")
        out("這一列已經被改動，但沒有存檔。")
        out("如果這只是測試，請重新查詢一次，提示選「否」還原。")
    else:
        out("全部都失敗。這一格目前的值：")
        try:
            out(f"  {cell.iface_value.CurrentValue!r}")
        except Exception:
            pass
        out("請重新查詢一次，提示選「否」還原。")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f"\n>>> 已存成 {OUT}")


if __name__ == "__main__":
    main()
