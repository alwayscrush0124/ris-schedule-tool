"""從傾印檔產生各院區的設定（帳號清單 + 對照表骨架）。

產出：
    設定/<院區>/accounts.csv   代班帳號清單（GUI 的「改成」下拉會讀它）
    設定/<院區>/mapping.csv    對照表骨架（規則待補，補之前用手動模式）

資料來源與可靠度：
  1. 傾印_使用者設定檔_<院區>_B.csv
     帳號和姓名是分開的兩欄 -> 完全不用推測，最可靠。
     班表上顯示的字串 = 帳號 + 姓名。
  2. 傾印_醫師週班設定_<院區>.csv
     公版班表實際用到的排班對象。裡面有「班表群組」（例如 NEUCTOFF均分、
     R01AM均分），那些不是使用者帳號，所以來源 1 找不到。
     這種只知道顯示名稱，不知道「要打什麼字進去」——
     去掉結尾的「均分」是常見情況但不是通則
     （W1PLAIN + W1X光均分 就不成立），所以一律標成未驗證。

     不過猜錯是安全的：程式打完字會比對下拉補完的結果，
     不符就按 ESC 取消並中止，不可能寫錯進去。

重跑會覆蓋 accounts.csv，但**不會動已經有內容的 mapping.csv**。
"""

import csv
import glob
import os
import re

CONFIG_DIR = "設定"

# 舊的手工帳號清單。裡面有實際存檔驗證過的帳號，價值比傾印檔推導的高，
# 一律優先採用 —— 例如 0生醫公用帳號 這一筆，
# 從傾印檔只推得出「整串當帳號」（錯的），實測要打的是 0生醫。
LEGACY_ACCOUNTS = "accounts.csv"
LEGACY_CAMPUS = "新竹臺大分院(含生醫醫院)"

# 這幾個小院區沒有自己的「醫師班表維護」排班醫師下拉，
# 實測（2026-07-30，重新查詢後再掃還是一樣）是直接沿用台大總院那份。
# 不是傾印時資料沒刷新的舊雷，是真的共用，所以不必各自存一份重複的白名單檔。
SHARED_USERCODE = {
    "北護分院": "台大總院",
    "金山分院": "台大總院",
}

# 代班／均分／公用這類「不是某個真人」的排班對象
POOL_HINTS = ("均分", "公用", "OFF")

# 這些也不是真人，但屬於公版的固定排班，通常不是請假要換的對象
STANDING_HINTS = ("未提成醫令", "國健", "抽審", "報告醫師", "X光", "LOCK")

# 除了排班醫師以外，可以改的欄位（排班醫師走 accounts/doctors 那條路）
OPTION_COLS = ["排班時段", "統計用分類碼", "排班類別",
               "檢查室名稱", "門急住限制", "檢查地點位址"]

# 逐一點過檢查地點倒出來的檢查室清單放這裡（dump_rooms.py 產生）
ROOMS_DIR = "檢查室"

# 代碼檔維護倒出來的代碼表。這幾欄是全院共用的，不分院區。
CODE_FILES = {
    "排班類別": "傾印_代碼_排班類別.csv",
    "門急住限制": "傾印_代碼_門急住限制.csv",
}

# 公版班表的欄名 -> 每日班表的欄名。差一個字，對不齊就整欄讀成空的。
WEEKLY_TO_DAILY = {
    "排班時段": "排班時段",
    "統計用分類碼": "統計用分類碼",
    "排班類別": "排班類別",
    "檢查室": "檢查室名稱",
    "門急住限制": "門急住限制",
    "檢查地點位置": "檢查地點位址",
}


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def is_pool(text):
    return any(h in text for h in POOL_HINTS)


def is_standing(text):
    return any(h in text for h in STANDING_HINTS)


def campus_from(path, prefix):
    name = os.path.basename(path).replace(prefix, "").replace(".csv", "")
    return name.replace("_B", "").replace("_D", "")


def load_accounts_dump():
    """{院區: [(帳號, 姓名), ...]}"""
    out = {}
    for f in glob.glob("傾印_使用者設定檔_*_B.csv"):
        camp = campus_from(f, "傾印_使用者設定檔_")
        rows = []
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            code, name = norm(r.get("使用者帳號")), norm(r.get("姓名"))
            if code and code != "(null)":
                rows.append((code, name))
        out[camp] = rows
    return out


def load_usercode_dump():
    """{院區: [顯示文字, ...]}

    「醫師班表維護」排班醫師下拉選單的全部選項（4_dump_options.py 產生）。
    這是 RIS 自己認定合法的排班醫師值，比使用者設定檔的全部帳號乾淨——
    不含 LOCK、測試帳號那些雜項。目前只有新竹臺大分院掃過。
    """
    out = {}
    for f in glob.glob("傾印_排班醫師選項_*.csv"):
        camp = campus_from(f, "傾印_排班醫師選項_")
        items = []
        with open(f, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.rstrip("\n").rstrip("\r")
                if not line.strip() or line.lstrip().startswith("#"):
                    continue
                if line == "顯示文字":
                    continue
                items.append(line)
        out[camp] = items
    return out


def filter_by_usercode(accounts, usercode_items):
    """把 accounts（來自使用者設定檔，帳號和姓名分開）按下拉選項白名單縮減。

    下拉選單顯示的文字有時是「帳號+姓名」，有時只有姓名甚至姓名前面多一個
    不相干的字元（實測踩過，例如 BIHRAD 顯示成「0生醫公用帳號」）——
    所以比對規則按可信度分三層，每層都要求唯一命中才採用：
      1. 帳號+姓名 完全相符 -> 帳號（使用者設定檔那份）就是對的，照抄
      2. 姓名完全相符（下拉選項本身沒有另外的帳號前綴）-> 真正要打的值
         就是姓名本身，不是使用者設定檔那個帳號
      3. 顯示文字以姓名結尾（唯一時才採用，姓名重複如「均分」不會走到這層）
         -> 真正的帳號是該選項扣掉姓名之後剩下的前綴（例如「WYC*王彧宸」
         的帳號是 WYC*），不是使用者設定檔那個帳號
    2026-08-10 踩到的坑：雲林分院登入帳號（WYCWYC／HJHHJH／Y00049…）
    和「排班醫師」下拉真正吃的帳號（WYC*／HJH*…）是兩套系統，落到 tier
    2/3 時如果照抄原本的 code，打進格子會被下拉補完成別的選項、驗證失敗、
    整批中止。所以 tier 2/3 一律回傳「從下拉選項反推出來的」帳號 + 顯示文字，
    不能沿用 accounts 原本的 code。
    三層都對不上、或命中不只一筆，保守起見不收（避免用猜的帳號寫錯格子）。
    """
    keep = []
    for code, name in accounts:
        display = code + name
        if any(it.strip() == display for it in usercode_items):
            keep.append((code, display))
            continue
        byname = [it for it in usercode_items if it.strip() == name]
        if len(byname) == 1:
            keep.append((name, name))
            continue
        end = [it for it in usercode_items
                if name and it.strip().endswith(name)]
        if len(end) == 1:
            opt = end[0].strip()
            real_code = opt[:-len(name)] if name else opt
            keep.append((real_code or name, opt))
    return keep


def load_weekly_dump():
    """{院區: Counter(排班醫師字串)}"""
    out = {}
    for f in glob.glob("傾印_醫師週班設定_*.csv"):
        camp = campus_from(f, "傾印_醫師週班設定_")
        seen = {}
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            v = norm(r.get("排班醫師"))
            if v and v != "(null)":
                seen[v] = seen.get(v, 0) + 1
        out[camp] = seen
    return out


def guess_code(display, known_codes):
    """從顯示名稱推「要打什麼字」。回傳 (帳號, 說明)。

    只在來源 1 找不到時才用，一律標未驗證。
    """
    for code in sorted(known_codes, key=len, reverse=True):
        if display.startswith(code):
            return code, "帳號取自使用者設定檔"
    if display.endswith("均分"):
        return display[:-2], "推測：去掉結尾的「均分」，第一次用請只改一列"
    return display, "推測：直接用顯示名稱，第一次用請只改一列"


def collect_options(camp):
    """這個院區各欄位實際用得到哪些值。回傳 {欄位: {值: 出現次數}}

    來源一：公版班表。它是母體 —— 該院區所有班別組合都在裡面。
            欄名跟每日班表差一點點（檢查室/檢查室名稱、檢查地點位置/位址），
            要轉成每日班表的叫法，因為工具是在每日班表上操作。
    來源二：工作時段設定。班表上的「排班時段」顯示的是那裡的「時段簡述」，
            全部 48 種都列出來，不必等它出現在某天的班表上才知道。
    """
    out = {c: {} for c in OPTION_COLS}

    for f in glob.glob("傾印_醫師週班設定_*.csv"):
        if campus_from(f, "傾印_醫師週班設定_") != camp:
            continue
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            for src, col in WEEKLY_TO_DAILY.items():
                v = norm(r.get(src))
                if v and v != "(null)":
                    out[col][v] = out[col].get(v, 0) + 1

    for f in glob.glob("傾印_工作時段設定*.csv"):
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            v = norm(r.get("時段簡述"))
            if v and v != "(null)":
                out["排班時段"].setdefault(v, 0)

    # 來源三：檢查地點及檢查室設定。公版只看得到「有排班的檢查室」，
    #         這裡是系統裡登記的全部，而且帶院區別可以分院區。
    for f in glob.glob(os.path.join(ROOMS_DIR, "*.csv")):
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            if norm(r.get("院區別")) != camp:
                continue
            v = norm(r.get("檢查室名稱"))
            if v and v != "(null)":
                out["檢查室名稱"].setdefault(v, 0)
            v = norm(r.get("檢查地點名稱"))
            if v and v != "(null)":
                out["檢查地點位址"].setdefault(v, 0)

    # 來源四：代碼檔維護。這幾欄是全院共用的代碼表，不分院區。
    for col, f in CODE_FILES.items():
        if not os.path.exists(f):
            continue
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            v = norm(r.get("代碼說明"))
            if v and v != "(null)" and norm(r.get("啟用否")) != "False":
                out[col].setdefault(v, 0)
    return out


def load_codes():
    """{欄位: {代碼說明: 代碼}}

    代碼檔維護裡代碼和顯示是分開的（B -> Body、O -> 門診），
    跟帳號/顯示名稱是同一個結構。班表上顯示的是「代碼說明」。
    要打進去的到底是代碼還是說明，還沒實測過，所以先只當參考資訊記著。
    """
    out = {}
    for col, f in CODE_FILES.items():
        if not os.path.exists(f):
            continue
        m = {}
        for r in csv.DictReader(open(f, encoding="utf-8-sig")):
            desc, code = norm(r.get("代碼說明")), norm(r.get("代碼"))
            if desc and desc != "(null)":
                m[desc] = code
        out[col] = m
    return out


def load_legacy():
    """讀舊的手工帳號清單，以「顯示名稱」為鍵。回傳 {顯示名稱: (帳號, 已驗證, 說明)}"""
    out = {}
    if not os.path.exists(LEGACY_ACCOUNTS):
        return out
    with open(LEGACY_ACCOUNTS, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            p = [x.strip() for x in next(csv.reader([line]))]
            if not p or p[0] == "帳號":
                continue
            while len(p) < 4:
                p.append("")
            out[p[1] or p[0]] = (p[0], p[2], p[3])
    return out


def build_campus(camp, accounts, weekly, legacy, usercode=None, shared_from=None):
    folder = os.path.join(CONFIG_DIR, camp)
    os.makedirs(folder, exist_ok=True)

    known = {c for c, _n in accounts}
    entries = []      # (帳號, 顯示名稱, 已驗證, 說明, 排序權重)
    seen = set()

    # 來源 1：使用者設定檔，帳號姓名分開，最可靠
    for code, name in accounts:
        display = code + name
        if not (is_pool(code + name)):
            continue
        entries.append((code, display, "N", "來自使用者設定檔（帳號可靠）", 0))
        seen.add(display)

    # 來源 2：公版班表實際用到的非真人排班對象
    for display, n in sorted(weekly.items(), key=lambda x: -x[1]):
        if display in seen:
            continue
        if not (is_pool(display) or is_standing(display)):
            continue
        code, why = guess_code(display, known)
        rank = 1 if is_pool(display) else 2
        note = f"公版用到 {n} 列；{why}"
        if rank == 2:
            note = f"公版固定排班（{n} 列），通常不是代班對象；{why}"
        entries.append((code, display, "N", note, rank))
        seen.add(display)

    # 手工清單優先：它是實際打進 RIS、存檔後重新查詢驗證過的，
    # 比從傾印檔推導的可靠。同一個顯示名稱就用手工那筆的帳號與驗證狀態。
    merged = []
    for code, display, ok, note, rank in entries:
        if display in legacy:
            lcode, lok, lnote = legacy[display]
            merged.append((lcode, display, lok or "N",
                           f"{lnote}（人工清單優先）", rank))
        else:
            merged.append((code, display, ok, note, rank))
    entries = merged

    # 手工清單裡有、但傾印檔沒出現的，也要留著
    have = {d for _c, d, _o, _n, _r in entries}
    for display, (lcode, lok, lnote) in legacy.items():
        if display not in have:
            entries.append((lcode, display, lok or "N",
                            f"{lnote}（只在人工清單裡）", 3))

    entries.sort(key=lambda e: (e[4], e[0]))

    path = os.path.join(folder, "accounts.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(f"# {camp} 代班帳號清單 —— GUI 的「改成」下拉選單會讀這個檔\n")
        f.write("#\n")
        f.write("# 帳號 = 實際打進格子的值；顯示名稱 = 打完之後畫面會顯示的樣子\n")
        f.write("# 已驗證 = Y 表示存檔後重新查詢確認過；N 表示還沒。\n")
        f.write("# 就算帳號猜錯也不會寫錯：程式打完會比對下拉補完的結果，\n")
        f.write("# 不符就按 ESC 取消並中止。但第一次用還是請勾「先只改一列試試」。\n")
        w = csv.writer(f)
        w.writerow(["帳號", "顯示名稱", "已驗證", "說明"])
        for code, display, ok, note, _r in entries:
            w.writerow([code, display, ok, note])

    # 醫師清單：該院區的所有帳號。
    # 代班不一定是丟給均分池，也可能直接指定某位同事接手，
    # 所以下拉選單要看得到人，不能只有池帳號。
    #
    # 有排班醫師下拉白名單（傾印_排班醫師選項_<院區>.csv）的院區，
    # 用它篩掉 LOCK / 測試 / 非醫師的雜項帳號——那是 RIS 自己認定合法的
    # 排班醫師值，比使用者設定檔的全部帳號乾淨。沒有白名單就照舊全部列出。
    # (code, name) 用在沒有白名單的院區 -> 顯示名稱要現算 code+name；
    # 有白名單時 filter_by_usercode 已經反推出正確帳號，回傳的是
    # (code, display)，display 就是下拉選項原文，不能再現算一次。
    doctor_accounts = [(code, code + name) for code, name in accounts]
    filtered_note = "# GUI 的「改成」下拉會把代班帳號排前面、這份排後面。\n"
    if usercode:
        doctor_accounts = filter_by_usercode(accounts, usercode)
        if shared_from:
            filtered_note = (
                f"# 這個院區沒有自己的排班醫師下拉，實測是共用「{shared_from}」那份\n"
                f"# （見 傾印_排班醫師選項_{shared_from}.csv），已用它篩過。\n"
                "# GUI 的「改成」下拉會把代班帳號排前面、這份排後面。\n"
            )
        else:
            filtered_note = (
                "# 已用「醫師班表維護」排班醫師下拉選單過濾，只留 RIS 認定合法的值\n"
                "# （傾印_排班醫師選項_<院區>.csv）；LOCK / 測試 / 非醫師帳號不在這裡。\n"
                "# GUI 的「改成」下拉會把代班帳號排前面、這份排後面。\n"
            )

    # tier2/3 反推帳號時，同一個人若在使用者設定檔裡有不只一個舊帳號
    # （實測雲林分院多人如此），會反推出同一個 (code, display)，去重。
    seen = set()
    doctor_accounts_dedup = []
    for code, display in doctor_accounts:
        key = (code, display)
        if key in seen:
            continue
        seen.add(key)
        doctor_accounts_dedup.append(key)

    dpath = os.path.join(folder, "doctors.csv")
    with open(dpath, "w", encoding="utf-8-sig", newline="") as f:
        f.write(f"# {camp} 全部帳號（來自使用者設定檔，帳號與姓名是分開的兩欄）\n")
        f.write(filtered_note)
        w = csv.writer(f)
        w.writerow(["帳號", "顯示名稱"])
        for code, display in doctor_accounts_dedup:
            w.writerow([code, display])

    # 欄位選項：各欄位可以填什麼。
    # 來源是公版班表（該院區實際用到的值）＋工作時段設定（全部 48 種時段），
    # 比「目前畫面上出現過的值」完整得多 —— 畫面只查得到當天那幾列。
    opts = collect_options(camp)
    codes = load_codes()
    opath = os.path.join(folder, "欄位選項.csv")
    with open(opath, "w", encoding="utf-8-sig", newline="") as f:
        f.write(f"# {camp} 各欄位可以填什麼 —— GUI 的「改成」下拉會讀這個檔\n")
        f.write("# 來源：公版班表（該院區實際用到的值）+ 工作時段設定（全部時段）\n")
        f.write("# 只是選項清單，不代表每個值在每一列都合法。\n")
        w = csv.writer(f)
        w.writerow(["欄位", "值", "公版出現次數", "代碼"])
        for col in OPTION_COLS:
            for val, n in sorted(opts.get(col, {}).items(),
                                 key=lambda x: (-x[1], x[0])):
                w.writerow([col, val, n, codes.get(col, {}).get(val, "")])

    # 對照表骨架：已經有內容就不動，免得蓋掉人工維護的規則
    mpath = os.path.join(folder, "mapping.csv")
    if not os.path.exists(mpath):
        with open(mpath, "w", encoding="utf-8-sig", newline="") as f:
            f.write(f"# {camp} 休假代班對照表\n")
            f.write("#\n")
            f.write("# 規則還沒建立。在補完之前，請用 GUI 的「全部改成」手動模式：\n")
            f.write("# 自己篩選出要改的列，自己選帳號。\n")
            f.write("#\n")
            f.write("# 下面這行宣告「拿哪幾欄來比對」，各院區可以不一樣。\n")
            f.write("# 例如有的院區看檢查項目，有的院區看星期和時段。\n")
            f.write("#!比對欄位=統計用分類碼,排班類別,檢查室名稱\n")
            f.write("#\n")
            f.write("# 星號 * 代表這一欄不參與比對；由上往下比對，第一個對上的就採用。\n")
            f.write("# 代班帳號留空 = 不自動處理，列進「需手動」。\n")
            w = csv.writer(f)
            w.writerow(["統計用分類碼", "排班類別", "檢查室名稱",
                        "代班帳號", "顯示名稱", "備註"])

    return len(entries), path


def main():
    accounts = load_accounts_dump()
    weekly = load_weekly_dump()
    legacy_all = load_legacy()
    usercode_all = load_usercode_dump()
    camps = sorted(set(accounts) | set(weekly))
    if not camps:
        raise SystemExit("找不到傾印檔，請先跑 dump_screen.py")

    print(f"{'院區':<28}{'帳號清單':>8}{'公版':>8}{'產生':>6}")
    for camp in camps:
        acc = accounts.get(camp, [])
        wk = weekly.get(camp, {})
        legacy = legacy_all if camp == LEGACY_CAMPUS else {}
        shared_from = SHARED_USERCODE.get(camp) if camp not in usercode_all else None
        usercode = usercode_all.get(camp) or usercode_all.get(shared_from)
        n, path = build_campus(camp, acc, wk, legacy, usercode, shared_from)
        print(f"{camp:<28}{len(acc):>8}{len(wk):>8}{n:>6}")
    print(f"\n設定寫在 {CONFIG_DIR}/ 底下")


if __name__ == "__main__":
    main()
