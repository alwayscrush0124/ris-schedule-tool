"""
RIS 醫師班表維護 - 視窗結構探測工具
=====================================
這支程式是「唯讀」的：它只讀取畫面上有哪些欄位、按鈕、表格，
不會點擊、不會輸入、不會存檔、不會連資料庫。

用法：
  1. 先手動開啟 RIS，進到「醫師班表維護」畫面
  2. 隨便查詢一次，讓表格裡有資料（這樣才看得到表格結構）
  3. 執行： python 1_inspect.py
  4. 結果會存成 window_dump.txt，把它給我看
"""

import sys
import io
from datetime import datetime

try:
    from pywinauto import Desktop
except ImportError:
    print("!! 還沒安裝 pywinauto，請先執行： pip install pywinauto")
    sys.exit(1)

OUT = "window_dump.txt"
# 用來認出目標視窗的關鍵字（找不到時會列出所有視窗讓你挑）
KEYWORDS = ["班表", "醫師", "RIS", "Unimax", "工作表"]


def list_all_windows():
    """列出目前所有看得到的視窗標題。"""
    rows = []
    for backend in ("uia", "win32"):
        try:
            for w in Desktop(backend=backend).windows():
                try:
                    title = w.window_text().strip()
                    cls = w.class_name()
                except Exception:
                    continue
                if title:
                    rows.append((backend, title, cls))
        except Exception as e:
            rows.append((backend, f"<列舉失敗: {e}>", ""))
    return rows


def main():
    buf = io.StringIO()

    def out(s=""):
        print(s)
        buf.write(str(s) + "\n")

    out(f"=== RIS 視窗探測 {datetime.now():%Y-%m-%d %H:%M:%S} ===\n")

    windows = list_all_windows()
    out("--- 目前所有視窗 ---")
    for backend, title, cls in windows:
        out(f"[{backend:5}] {title}   (class={cls})")

    # 挑出可能是目標的視窗
    hits = [
        (b, t, c) for b, t, c in windows
        if any(k in t for k in KEYWORDS)
    ]

    out("\n--- 可能是目標的視窗 ---")
    if not hits:
        out("找不到。請確認 RIS 的「醫師班表維護」畫面是開著的，然後重跑一次。")
        out("如果畫面明明開著卻找不到，把上面那份完整視窗清單給我看就好。")
    for b, t, c in hits:
        out(f"[{b:5}] {t}")

    # 對每個命中的視窗，倒出完整的控制項結構
    for backend, title, _cls in hits:
        out(f"\n\n=========== 控制項結構: [{backend}] {title} ===========")
        try:
            win = Desktop(backend=backend).window(title=title)
            dump = io.StringIO()
            old = sys.stdout
            sys.stdout = dump
            try:
                # depth 限制避免輸出爆炸；不夠深再調大
                win.print_control_identifiers(depth=6)
            finally:
                sys.stdout = old
            out(dump.getvalue())
        except Exception as e:
            out(f"（讀取失敗: {e}）")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(buf.getvalue())

    print(f"\n>>> 已存成 {OUT}，把這個檔案給我看。")


if __name__ == "__main__":
    main()
