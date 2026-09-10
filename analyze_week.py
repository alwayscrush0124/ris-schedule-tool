"""看懂「醫師週班設定」這份公版班表。

這份是公版：沒有人休假時，展開之後就照它執行。
所以它要回答的問題是「平常誰上什麼班」，
而不是「休假要換成哪個 OFF 帳號」—— 後者是 mapping.csv 的事，
兩者不要混在一起看。

特別要留意「有時段但沒有醫師」的列：那是刻意留白的，
因為輪值規則不固定，展開之後要人工補上。

結果寫進 分析_週班表.txt（主控台是 cp950，中文直接印會亂碼）。
"""

import csv
import glob
import io
from collections import Counter, defaultdict

SRC_HINT = "週班"
BLANK = ("", "(null)", "不限")

# 週班設定叫「檢查地點位置」，每日班表叫「檢查地點位址」——
# 只差一個字，讀錯就會以為整欄是空的。兩個都試。
LOCATION_COLS = ("檢查地點位置", "檢查地點位址")


def norm(s):
    return str(s or "").strip()


def main():
    files = [x for x in glob.glob("傾印_*.csv") if SRC_HINT in x]
    if not files:
        raise SystemExit("找不到週班設定傾印檔，請先跑 dump_screen.py")
    src = files[0]
    rows = list(csv.DictReader(open(src, encoding="utf-8-sig")))

    # 檔名帶院區，各院區的分析才不會互相蓋掉
    out_name = "分析_" + src.replace("傾印_", "").replace(".csv", "") + ".txt"

    def loc(r):
        for c in LOCATION_COLS:
            if c in r:
                return norm(r[c])
        return ""

    out = io.StringIO()

    def p(s=""):
        out.write(str(s) + "\n")

    p(f"來源：{src}")
    p(f"公版班表 {len(rows)} 列")
    p()

    # ---- 有沒有排到人 ----
    filled = [r for r in rows if norm(r["排班醫師"]) not in ("", "(null)")]
    empty = [r for r in rows if norm(r["排班醫師"]) in ("", "(null)")]
    p(f"已排定醫師 {len(filled)} 列")
    p(f"留白待補   {len(empty)} 列   <- 輪值不固定，展開後要人工補")
    p()

    if empty:
        p("=" * 96)
        p("留白的班（展開後要自己填）")
        p("=" * 96)
        p(f"{'星期':<6}{'分類':<8}{'類別':<20}{'時段':<24}{'檢查室':<20}")
        for r in sorted(empty, key=lambda x: (x["星期"], x["統計用分類碼"],
                                              x["排班類別"])):
            p(f"{norm(r['星期']):<6}{norm(r['統計用分類碼']):<8}"
              f"{norm(r['排班類別']):<20}{norm(r['排班時段']):<24}"
              f"{norm(r['檢查室']):<20}")
        p()

    # ---- 每天幾列 ----
    p("=" * 96)
    p("每個星期幾的班數")
    p("=" * 96)
    per_day = Counter(norm(r["星期"]) for r in rows)
    for d, n in sorted(per_day.items()):
        blank_n = sum(1 for r in empty if norm(r["星期"]) == d)
        p(f"  {d:<8}{n:>4} 列" + (f"（其中 {blank_n} 列留白）" if blank_n else ""))
    p()

    # ---- 誰的班最多 ----
    p("=" * 96)
    p("每位醫師的班數")
    p("=" * 96)
    per_doc = Counter(norm(r["排班醫師"]) for r in filled)
    for doc, n in per_doc.most_common():
        p(f"  {doc:<20}{n:>4} 列")
    p()

    # ---- 班別組合 ----
    p("=" * 96)
    p("班別組合（分類 / 類別 / 時段 / 檢查室）× 一週出現幾天")
    p("=" * 96)
    combo = defaultdict(set)
    combo_docs = defaultdict(set)
    for r in rows:
        key = (norm(r["統計用分類碼"]), norm(r["排班類別"]),
               norm(r["排班時段"]), norm(r["檢查室"]))
        combo[key].add(norm(r["星期"]))
        d = norm(r["排班醫師"])
        combo_docs[key].add(d if d not in ("", "(null)") else "(留白)")
    p(f"共 {len(combo)} 種")
    p()
    p(f"{'分類':<8}{'類別':<20}{'時段':<24}{'檢查室':<20}{'天':>3}  固定的人")
    for key in sorted(combo, key=lambda k: (-len(combo[k]), k)):
        docs = sorted(combo_docs[key])
        who = "、".join(docs) if len(docs) <= 3 else f"{len(docs)} 人輪"
        p(f"{key[0]:<8}{key[1]:<20}{key[2]:<24}{key[3]:<20}"
          f"{len(combo[key]):>3}  {who}")

    # 地點分布：竹東那些班在每日班表會被排除規則擋掉，先看看有多少
    p()
    p("=" * 96)
    p("檢查地點位置分布")
    p("=" * 96)
    for k, n in Counter(loc(r) for r in rows).most_common():
        p(f"  {k or '(空白)':<12}{n:>4} 列")

    with open(out_name, "w", encoding="utf-8") as f:
        f.write(out.getvalue())
    print(f"OK -> {out_name}")


if __name__ == "__main__":
    main()
