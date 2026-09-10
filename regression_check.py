"""回歸驗證：院區化改版前後，對同一批資料的判定必須完全一致。

改版動了 load_mapping / match_rule / plan_for 的結構
（比對欄位改成由設定檔宣告、竹東排除規則搬進設定檔）。
這些是決定「哪一列要改成什麼」的核心，改壞了會直接寫錯資料，
所以不能只靠「看起來沒問題」。

作法：
    舊版程式碼備份在 _ris_gui_before.py，
    拿它和新版各跑一次判定，逐列比對。
    資料用公版班表（367 列，涵蓋 55 種班別組合），欄名轉成每日班表的樣子。

不寫入 RIS，純計算。
"""

import csv
import glob
import importlib.util
import sys

# 週班設定的欄名跟每日班表差一點點，要對齊才能餵進判定邏輯
RENAME = {"檢查室": "檢查室名稱", "檢查地點位置": "檢查地點位址"}


def load_module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def rows_from_weekly():
    files = [x for x in glob.glob("傾印_醫師週班設定_*.csv") if "新竹臺大" in x]
    if not files:
        raise SystemExit("找不到新竹臺大的公版班表傾印檔")
    out = []
    for r in csv.DictReader(open(files[0], encoding="utf-8-sig")):
        row = {RENAME.get(k, k): v for k, v in r.items()}
        out.append(row)
    return out


def main():
    old = load_module("_ris_gui_before.py", "ris_gui_old")
    new = load_module("ris_gui.py", "ris_gui_new")

    old_rules = old.load_mapping("mapping.csv")
    new_rules = new.load_mapping("設定/新竹臺大分院(含生醫醫院)/mapping.csv")
    print(f"舊版規則 {len(old_rules)} 條，新版規則 {len(new_rules)} 條")
    print(f"新版比對欄位 {new_rules.match_cols}")
    print(f"新版排除     {new_rules.excludes}")

    rows = rows_from_weekly()
    print(f"驗證資料 {len(rows)} 列\n")

    diff = 0
    for i, row in enumerate(rows):
        a = old.plan_for(row, old_rules)
        b = new.plan_for(row, new_rules)
        # 說明文字允許不同（竹東那條的措辭有改），帳號和顯示名稱必須一致
        if a[0] != b[0] or a[1] != b[1]:
            diff += 1
            if diff <= 10:
                print(f"  !! 第 {i} 列不一致")
                print(f"     {row.get('統計用分類碼')} / {row.get('排班類別')} / "
                      f"{row.get('檢查室名稱')} / {row.get('檢查地點位址')}")
                print(f"     舊: {a}")
                print(f"     新: {b}")

    if diff:
        print(f"\n*** 有 {diff} 列判定不一致，改版改壞了東西 ***")
        sys.exit(1)
    print(f"通過：{len(rows)} 列的判定完全一致（帳號與顯示名稱）")

    # 說明文字的差異單獨列出來看一眼，確認只有預期中的措辭變動
    notes = {}
    for row in rows:
        a = old.plan_for(row, old_rules)
        b = new.plan_for(row, new_rules)
        if a[2] != b[2]:
            notes[(a[2], b[2])] = notes.get((a[2], b[2]), 0) + 1
    if notes:
        print("\n說明文字有變動的（不影響行為）：")
        for (x, y), n in notes.items():
            print(f"  {n:>4} 列   {x!r} -> {y!r}")


if __name__ == "__main__":
    main()
