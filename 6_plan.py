"""
RIS 休假代班 - 產生異動計畫（只看不改）
========================================
讀取你已經查詢出來的表格，比對 mapping.csv，
算出每一列該改成哪個 OFF 帳號，然後印出報告。

*** 這支程式不會修改任何東西。***
不點擊、不輸入、不存檔。只讀畫面 + 算 + 印報告。

用法：
  1. 在 RIS 裡照平常的方式查詢（填日期 + 排班醫師）
  2. python 6_plan.py G00453
     只改半天： python 6_plan.py G00453 --shift AM
                python 6_plan.py G00453 --shift PM
  3. 看報告。確認沒問題後，再跑 7_apply.py 實際修改。
"""

import argparse
import csv
import io
import os
import re
import sys

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"
MAPPING = "mapping.csv"
PLAN_OUT = "plan.csv"

COLS = ["起始日期", "起始時間", "結束日期", "結束時間", "排班時段",
        "檢查室名稱", "門急住限制", "統計用分類碼", "排班類別",
        "檢查地點位址", "排班醫師"]

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def norm(s):
    """比對前正規化：去頭尾空白、全形空白。"""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def load_mapping(path):
    """讀對照表。只讀主區段（統計用分類碼開頭那段），忽略均分池區段。"""
    rules = []
    if not os.path.exists(path):
        out(f"!! 找不到 {path}")
        return rules

    in_main = False
    with open(path, encoding="utf-8-sig") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = next(csv.reader([line]))
            parts = [p.strip() for p in parts]
            if parts[0] == "統計用分類碼":
                in_main = True
                continue
            if parts[0] == "池帳號":
                in_main = False
                continue
            if not in_main or len(parts) < 4:
                continue
            rules.append({
                "統計用分類碼": norm(parts[0]),
                "排班類別": norm(parts[1]),
                "檢查室名稱": norm(parts[2]),
                "代班帳號": norm(parts[3]),
                "顯示名稱": norm(parts[4]) if len(parts) > 4 else "",
                "備註": norm(parts[5]) if len(parts) > 5 else "",
            })
    return rules


def match_rule(row, rules):
    """由上往下找第一個對上的規則。回傳 (代班帳號, 規則) 或 (None, None)。"""
    for r in rules:
        for key in ("統計用分類碼", "排班類別", "檢查室名稱"):
            want = r[key]
            if want == "*":
                continue
            if norm(row.get(key)) != want:
                break
        else:
            return (r["代班帳號"] or None), r
    return None, None


def cell_value(cell):
    try:
        v = cell.iface_value.CurrentValue
        if v not in (None, ""):
            return v
    except Exception:
        pass
    try:
        v = cell.legacy_properties().get("Value")
        if v not in (None, ""):
            return v
    except Exception:
        pass
    return cell.window_text()


def read_grid():
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
        sys.exit(1)

    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    win = app.window(handle=hwnd)
    form = win.child_window(auto_id=FORM_AUTO_ID)
    grid = form.child_window(auto_id=GRID_AUTO_ID)
    grid.wait("exists", timeout=15)

    # 順便把畫面上的查詢條件記下來，出問題時才看得出是查了什麼
    out("目前畫面上的查詢條件：")
    for label, aid in [("院區別", "HOSPITALCODE"), ("醫師類別", "USERLEVELCODE"),
                       ("排班醫師", "USERCODE"), ("起始日期", "STARTDATE"),
                       ("結束日期", "ENDDATE")]:
        try:
            ctl = form.child_window(auto_id=aid, control_type="ComboBox")
            out(f"    {label:6} = {ctl.window_text()!r}")
        except Exception:
            try:
                out(f"    {label:6} = {form.child_window(auto_id=aid).window_text()!r}")
            except Exception:
                out(f"    {label:6} = (讀不到)")
    out()

    children = grid.children()
    kinds = {}
    for c in children:
        try:
            n = c.window_text()
        except Exception:
            continue
        k = n.split()[0] if n.split() else "(空白)"
        kinds[k] = kinds.get(k, 0) + 1
    out(f"表格元件 {len(children)} 個，種類統計: {kinds}")

    rows = []
    for r in children:
        try:
            name = r.window_text()
        except Exception:
            continue
        if not name.startswith("資料列"):
            continue
        idx = int(name.split()[1])
        vals = []
        for c in r.children():
            t = c.window_text().split()
            if len(t) == 2 and t[0] == "資料列":
                continue
            vals.append(cell_value(c))
        if len(vals) < len(COLS):
            continue
        row = dict(zip(COLS, vals))
        if all(norm(v) in ("", "(null)") for v in row.values()):
            continue  # 最後那列空白的新增列
        row["_row"] = idx
        rows.append(row)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("doctor", help="休假醫師，可只打代號如 G00453")
    ap.add_argument("--shift", choices=["AM", "PM", "ALL"], default="ALL",
                    help="只處理上半天(AM)或下半天(PM)，預設全部")
    args = ap.parse_args()

    rules = load_mapping(MAPPING)
    out(f"對照表載入 {len(rules)} 條規則\n")

    rows = read_grid()
    out(f"表格讀到 {len(rows)} 列資料\n")

    if not rows:
        out("!! 表格是空的。請確認：")
        out("   1. 「醫師班表維護」畫面有開著")
        out("   2. 已經按過查詢，畫面上看得到資料")
        out("   3. 查詢條件的日期是你要處理的那一天")
        return

    # 表格裡有哪些醫師，方便對照名字有沒有打錯
    docs = {}
    for r in rows:
        d = norm(r["排班醫師"])
        docs[d] = docs.get(d, 0) + 1
    out("表格裡出現的排班醫師：")
    for d, n in sorted(docs.items(), key=lambda x: -x[1]):
        out(f"    {n:>3} 列  {d}")
    out()

    key = norm(args.doctor)
    plan, skip_other, manual, skip_zhudong, skip_shift = [], [], [], [], []

    for row in rows:
        doc = norm(row["排班醫師"])
        if key not in doc:
            skip_other.append(row)
            continue
        if norm(row["檢查地點位址"]) == "竹東":
            skip_zhudong.append(row)
            continue
        seg = norm(row["排班時段"])
        if args.shift == "AM" and "上半天" not in seg:
            skip_shift.append(row)
            continue
        if args.shift == "PM" and "下半天" not in seg:
            skip_shift.append(row)
            continue

        target, rule = match_rule(row, rules)
        if target:
            row["_target"] = target
            row["_display"] = rule.get("顯示名稱", "")
            plan.append(row)
        else:
            row["_why"] = (rule["備註"] if rule else "對照表裡沒有這個班別")
            manual.append(row)

    # ---------- 報告 ----------
    out("=" * 78)
    out(f"休假醫師: {args.doctor}    時段: {args.shift}")
    out("=" * 78)

    if not plan and not manual:
        out(f"\n!! 表格裡沒有任何一列的排班醫師含有 {args.doctor!r}。")
        out("   請對照上面那份醫師清單，確認代號打對了，")
        out("   或者這位醫師當天的班本來就掛在均分池帳號下（那要手動處理）。")

    out(f"\n■ 會修改的 {len(plan)} 列")
    if plan:
        out(f"  {'列':>3}  {'時段':22} {'分類':6} {'類別':14} {'目前':16} -> 改成")
        out("  " + "-" * 92)
        for r in plan:
            out(f"  {r['_row']:>3}  {norm(r['排班時段']):22} {norm(r['統計用分類碼']):6} "
                f"{norm(r['排班類別']):14} {norm(r['排班醫師']):16} -> "
                f"{r['_display'] or r['_target']}  (填入 {r['_target']})")

    if manual:
        out(f"\n■ 需要你手動處理的 {len(manual)} 列")
        for r in manual:
            out(f"  {r['_row']:>3}  {norm(r['統計用分類碼']):6} {norm(r['排班類別']):14} "
                f"{norm(r['檢查室名稱']):18} 原因: {r['_why']}")

    if skip_shift:
        out(f"\n■ 時段不符跳過 {len(skip_shift)} 列")
    if skip_zhudong:
        out(f"\n■ 竹東跳過 {len(skip_zhudong)} 列（竹東在別處設定）")
    if skip_other:
        out(f"\n■ 其他醫師的 {len(skip_other)} 列，不會動")

    out("\n" + "=" * 78)
    out("提醒：X 光那類掛在均分池帳號下的列，不會出現在這裡，要另外手動處理。")
    out("=" * 78)

    with open(PLAN_OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["表格列號"] + COLS + ["改成帳號", "改成顯示"])
        for r in plan:
            w.writerow([r["_row"]] + [r[c] for c in COLS]
                       + [r["_target"], r.get("_display", "")])
    out(f"\n異動計畫已存成 {PLAN_OUT}（給 7_apply.py 用）")

    with open("plan_report.txt", "w", encoding="utf-8") as f:
        f.write(_log.getvalue())


if __name__ == "__main__":
    main()
