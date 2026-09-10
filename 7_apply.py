"""
RIS 休假代班 - 實際修改（會改畫面，但絕不存檔）
================================================
讀 plan.csv，把表格裡「排班醫師」那一欄改成 OFF 帳號。

寫入的部分直接用 ris_gui.py 裡的 Ris 類別，不另外寫一份 ——
以前這支自己實作，結果用了會失敗的寫法（見下），修好一邊另一邊還是壞的。

安全設計：
  * 只碰「排班醫師」這一欄，其他 10 欄一個都不動。
  * 改每一列之前，會先重新核對這一列的內容跟計畫表一不一樣。
    只要有任何一格對不上就中止，不會硬改。
  * 用鍵盤打字寫入，打完先看下拉補完的結果對不對，
    不對就按 ESC 取消，那一列完全不會被動到。
  * 全部改完會重讀整張表逐列核對，連「不該動的列被動到」也會抓出來。
  * *** 全程不會按存檔。*** 改完停住，由你檢查後自己按。

用法：
  先跑 6_plan.py 產生 plan.csv，確認內容無誤，然後：
    python 7_apply.py --limit 1     # 強烈建議第一次只改一列
    python 7_apply.py               # 改全部

重要：執行後表格處於「已修改未存檔」狀態。
      這時不要關視窗、不要重新查詢，會跳出存檔提示。
      要放棄修改的話，在那個提示選「否」。
"""

import argparse
import csv
import io
from datetime import datetime

from ris_gui import DOCTOR_COL, Ris, norm

PLAN = "plan.csv"

# 核對用的欄位：這幾格必須跟計畫表完全一致，才允許修改
VERIFY_COLS = ["起始日期", "排班時段", "統計用分類碼", "排班類別",
               "檢查室名稱", "排班醫師"]

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0,
                    help="最多只改前幾列（0 = 全部）。第一次請用 --limit 1")
    args = ap.parse_args()

    with open(PLAN, encoding="utf-8-sig") as f:
        plan = list(csv.DictReader(f))
    if not plan:
        out("plan.csv 是空的，請先跑 6_plan.py。")
        return

    if args.limit:
        plan = plan[:args.limit]

    out("=" * 78)
    out(f"準備修改 {len(plan)} 列   {datetime.now():%Y-%m-%d %H:%M:%S}")
    out("=" * 78)
    for p in plan:
        out(f"  列{p['表格列號']:>3}  {norm(p['統計用分類碼']):6} {norm(p['排班類別']):14} "
            f"{norm(p['排班醫師']):16} -> {p.get('改成顯示') or p['改成帳號']}"
            f"  (填入 {p['改成帳號']})")

    print("\n" + "!" * 60)
    print("這會實際修改 RIS 畫面上的資料（但不會存檔）。")
    print("改完之後由你檢查，確認無誤再自己按存檔。")
    print("!" * 60)
    if input("確定要執行嗎？請完整輸入 yes： ").strip().lower() != "yes":
        out("\n已取消，什麼都沒改。")
        return

    ris = Ris()
    ris.connect()
    rows = ris.read_rows()
    out(f"\n表格讀到 {len(rows)} 列")
    snapshot = {r["_row"]: norm(r[DOCTOR_COL]) for r in rows}

    done, failed = 0, 0
    written = []
    attempted = set()   # 動過手的列（失敗的也算），收尾核對用

    for p in plan:
        idx = int(p["表格列號"])
        target = p["改成帳號"]            # 實際打進格子的值，例如 CCTAOFF
        display = p.get("改成顯示", "")    # 打完之後畫面該顯示的，例如 CCTAOFF均分
        out(f"\n--- 第 {idx} 列 ---")

        # 核對：這一列還是不是計畫表裡的那一列
        now = ris.read_row(idx)
        if now is None:
            out(f"  X 找不到第 {idx} 列，中止。")
            failed += 1
            break

        mismatch = [f"{c}: 畫面是 {norm(now.get(c))!r}，計畫表是 {norm(p.get(c, ''))!r}"
                    for c in VERIFY_COLS
                    if norm(now.get(c)) != norm(p.get(c, ""))]
        if mismatch:
            out("  X 內容跟計畫表對不上，中止以免改錯：")
            for m in mismatch:
                out(f"      {m}")
            out("  （表格可能被重新查詢過了。請重跑 6_plan.py。）")
            failed += 1
            break

        # 游標一定要先進到這一列，否則 RIS 不會認帳
        moved, how = ris.goto_row(idx)
        if not moved:
            out(f"  X {how}，中止（不在不確定的位置寫入）。")
            failed += 1
            break

        attempted.add(idx)
        ok, after, why = ris.write_cell(idx, DOCTOR_COL, target, display or None)
        if ok:
            out(f"  V {norm(p['排班醫師'])} -> {after}")
            written.append((idx, display or target))
            done += 1
        else:
            out(f"  X 打 {target} 的時候：{why}")
            out("     中止。請檢查畫面，需要的話重新查詢並選「否」還原。")
            failed += 1
            break

    # 收尾核對：重讀整張表，連沒打算動的列也要比對。
    # 只檢查目標列的話，「不該動的列被動到」這種錯永遠看不出來。
    if done:
        out("\n--- 重讀整張表核對 ---")
        want = dict(snapshot)
        want.update({i: norm(v) for i, v in written})
        try:
            fresh = ris.read_rows()
        except Exception as e:
            fresh = None
            out(f"  !! 重讀失敗（{type(e).__name__}），請自己在畫面上逐列確認。")

        if fresh is not None:
            wrong, strays = [], []
            for x in fresh:
                i, got = x["_row"], norm(x[DOCTOR_COL])
                if i not in want:
                    continue
                if got == want[i] or got.startswith(want[i]):
                    continue
                (wrong if i in attempted else strays).append((i, got, want[i]))

            for i, got, exp in wrong:
                out(f"  X 第 {i} 列：現在是 {got!r}，預期 {exp!r}")
            for i, got, exp in strays:
                out(f"  !! 第 {i} 列根本不該被動到，卻從 {exp!r} 變成 {got!r}")

            if strays:
                out("\n  *** 請不要存檔。重新查詢，提示選「否」還原。***")
            elif wrong:
                out("\n  *** 建議不要存檔，重新查詢並選「否」還原。***")
            else:
                out(f"  核對通過：{done} 列已改，其他列都沒被動到。")

    out("\n" + "=" * 78)
    out(f"完成 {done} 列" + (f"，失敗後中止（{failed}）" if failed else ""))
    out("=" * 78)
    if done:
        out("\n>>> 接下來請你自己做：")
        out("    1. 看一下畫面，確認改的內容正確")
        out("    2. 沒問題再按存檔")
        out("    3. 要放棄的話，重新查詢一次，跳出提示時選「否」")
        out("\n>>> 存檔後請重新查詢一次確認 —— 查不到資料才代表真的寫進去了。")
    out("\n程式沒有按存檔，目前資料庫裡還是原本的內容。")

    with open("apply_log.txt", "a", encoding="utf-8") as f:
        f.write(_log.getvalue() + "\n\n")
    print(">>> 紀錄已附加到 apply_log.txt")


if __name__ == "__main__":
    main()
