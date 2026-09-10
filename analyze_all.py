"""跨院區比對：各院區的公版班表用什麼東西當「排班對象」。

想回答的問題是：休假代班的邏輯在各院區是不是同一套？
如果每個院區的替代帳號體系不一樣，那對照表就不能只是換內容，
連「用哪幾欄決定換成什麼」都得跟著變。

作法：
  公版班表的「排班醫師」是「帳號+姓名」黏在一起（例如 G00453王小明）。
  我們有各院區權威的帳號/姓名清單（使用者設定檔傾印檔），
  所以可以精準拆開，不用猜 —— 這正是先前 CCTAOFF / CCTAOFF均分 一直搞錯的地方。

結果寫進 分析_跨院區.txt。
"""

import csv
import glob
import io
import re
from collections import Counter, defaultdict

OUT = "分析_跨院區.txt"

# 不是真人的排班對象，通常帶這些字樣
POOL_HINTS = ("均分", "公用", "OFF", "LOCK")


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def campus_from(path, prefix):
    """從檔名取院區。傾印_醫師週班設定_台大總院.csv -> 台大總院"""
    name = path.replace(prefix, "").replace(".csv", "")
    return name.replace("_B", "").replace("_D", "")


def load_accounts():
    """{院區: {帳號: 姓名}}"""
    out = {}
    for f in glob.glob("傾印_使用者設定檔_*_B.csv"):
        camp = campus_from(f, "傾印_使用者設定檔_")
        acc = {}
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            acc[norm(r.get("使用者帳號"))] = norm(r.get("姓名"))
        out[camp] = acc
    return out


def split_doctor(value, accounts):
    """把「帳號+姓名」拆回 (帳號, 姓名)。拆不出來回 (None, value)。

    用該院區的帳號清單去比對，取「開頭吻合而且最長」的那個帳號 ——
    最長是必要的：W1PLAIN 和 W1PLAINW1X光均分 同時存在時，
    取最短的會拆錯。
    """
    v = norm(value)
    best = None
    for code, name in accounts.items():
        if code and v.startswith(code):
            if best is None or len(code) > len(best):
                best = code
    if best is None:
        return None, v
    return best, v[len(best):]


def main():
    accounts = load_accounts()
    out = io.StringIO()

    def p(s=""):
        out.write(str(s) + "\n")

    p("各院區帳號清單：")
    for camp, acc in sorted(accounts.items()):
        p(f"  {camp:<28}{len(acc):>5} 個帳號")
    p()

    weekly = sorted(glob.glob("傾印_醫師週班設定_*.csv"))
    p("=" * 100)
    p("各院區公版班表裡的「非真人」排班對象")
    p("=" * 100)

    summary = {}
    for f in weekly:
        camp = campus_from(f, "傾印_醫師週班設定_")
        rows = list(csv.DictReader(open(f, encoding="utf-8-sig")))
        acc = accounts.get(camp, {})

        used = Counter()
        unknown = Counter()
        for r in rows:
            v = norm(r.get("排班醫師"))
            if v in ("", "(null)"):
                continue
            code, name = split_doctor(v, acc)
            if code is None:
                unknown[v] += 1
            else:
                used[(code, name)] += 1

        pools = {k: n for k, n in used.items()
                 if any(h in (k[0] + k[1]).upper() for h in
                        [h.upper() for h in POOL_HINTS])}
        people = {k: n for k, n in used.items() if k not in pools}

        summary[camp] = (len(rows), len(people), len(pools), sum(pools.values()))

        p()
        p(f"--- {camp} ---")
        p(f"    公版 {len(rows)} 列；真人 {len(people)} 位；"
          f"非真人 {len(pools)} 種（佔 {sum(pools.values())} 列）")
        if acc:
            for (code, name), n in sorted(pools.items(), key=lambda x: -x[1]):
                p(f"      {code:<22}{name:<18}{n:>4} 列")
        else:
            p("      （沒有這個院區的帳號清單，無法拆帳號）")
        if unknown:
            p(f"    拆不出帳號的 {len(unknown)} 種：")
            for v, n in unknown.most_common(8):
                p(f"      {v:<40}{n:>4} 列")

    p()
    p("=" * 100)
    p("摘要")
    p("=" * 100)
    p(f"{'院區':<28}{'公版列數':>8}{'真人':>6}{'非真人種類':>10}{'非真人列數':>10}")
    for camp, (a, b, c, d) in sorted(summary.items(), key=lambda x: -x[1][0]):
        p(f"{camp:<28}{a:>8}{b:>6}{c:>10}{d:>10}")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(out.getvalue())
    print(f"OK -> {OUT}")


if __name__ == "__main__":
    main()
