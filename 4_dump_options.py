"""
RIS 醫師班表維護 - 倒出下拉選單的所有選項
==========================================
目的：把「排班醫師」下拉選單裡的所有帳號撈出來，
      特別是各班別的 OFF 均分帳號（NEUCTOFF 之類）。
      順便撈「排班類別」「統計用分類」的完整選項，
      這樣之後比對才能用系統的原始寫法，不會因為打錯字對不上。

安全性：
  預設用「訊息查詢」的方式讀取，畫面上不會有任何變化，
  不會展開選單、不會改變目前選的值、不會存檔。
  萬一這個方式讀不到，才會問你要不要改用「展開選單」的方式。

用法：
  1. 開著「醫師班表維護」
  2. python 4_dump_options.py
  3. 把 options.txt 給我
"""

import io
import re

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
OUT = "options.txt"

# 院區別 (HOSPITALCODE) 就是這個畫面的查詢條件，讀它目前選中的文字
# 就知道現在開的是哪個院區，不用手動指定。讀不出來才退回這個預設值。
USERCODE_CAMPUS_FALLBACK = "新竹臺大分院(含生醫醫院)"
USERCODE_OUT_TMPL = "傾印_排班醫師選項_{campus}.csv"

TARGETS = [
    ("排班醫師", "USERCODE"),
    ("排班類別", "SCHEDULETYPECODE1"),
    ("統計用分類", "STATISTICGROUPCODE"),
    ("醫師類別", "USERLEVELCODE"),
    ("院區別", "HOSPITALCODE"),
]

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


def find_hwnd():
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                return w.handle, w.window_text()
        except Exception:
            continue
    return None, None


def dump_win32(hwnd):
    """用 win32 訊息直接問選單內容 —— 畫面完全不會動。"""
    app = Application(backend="win32").connect(handle=hwnd)
    main = app.window(handle=hwnd)
    got_any = False
    usercode_items = None
    campus = None

    for label, auto_id in TARGETS:
        out("\n" + "=" * 60)
        out(f"{label}  (auto_id={auto_id})")
        out("=" * 60)
        try:
            combo = main.child_window(auto_id=auto_id, control_type="System.Windows.Forms.ComboBox")
            items = combo.item_texts()
            if items:
                got_any = True
                out(f"共 {len(items)} 個選項：")
                for i, t in enumerate(items):
                    out(f"  [{i:3}] {t}")
                if auto_id == "USERCODE":
                    usercode_items = items
                if auto_id == "HOSPITALCODE":
                    try:
                        sel = combo.selected_text().strip()
                        if sel:
                            campus = sel
                            out(f"目前查詢院區：{sel}")
                    except Exception as e:
                        out(f"  讀目前選中的院區失敗: {type(e).__name__}: {e}")
            else:
                out("  （讀到 0 個選項）")
        except Exception as e:
            out(f"  讀取失敗: {type(e).__name__}: {e}")
    return got_any, usercode_items, campus


def save_usercode_csv(items, campus):
    """把「排班醫師」下拉選項存成 build_configs.py 認得的格式。

    這是 RIS 自己認定合法的排班醫師值，doctors.csv 會拿這份去過濾掉
    LOCK / 測試 / 非醫師的雜項帳號。
    """
    if not items:
        return
    if not campus:
        campus = USERCODE_CAMPUS_FALLBACK
        out(f"\n!! 讀不到目前查詢院區，先當成「{campus}」存檔，"
            f"存錯的話把檔名改掉重跑 build_configs.py。")
    path = USERCODE_OUT_TMPL.format(campus=campus)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(f"# {campus}「醫師班表維護」畫面排班醫師下拉選單的全部選項\n")
        f.write("# 來源：4_dump_options.py 讀 USERCODE 這個下拉（訊息查詢，畫面不會動）\n")
        f.write("# 這是 RIS 自己認定「可以填進排班醫師」的清單，用來把 doctors.csv 縮到只剩這些人\n")
        f.write("顯示文字\n")
        for t in items:
            if t.strip():
                f.write(t + "\n")
    out(f"\n>>> 排班醫師選項另存成 {path}")


def dump_uia_expand(hwnd):
    """備案：用 UIA 展開選單再讀。會看到選單彈出又收起。"""
    print("\n" + "=" * 60)
    print("備案模式：會把下拉選單「展開再收起」。")
    print("畫面上會看到選單閃一下，但不會改變目前選中的值、不會存檔。")
    print("=" * 60)
    if input("要執行嗎？請完整輸入 yes： ").strip().lower() != "yes":
        out("\n你選擇跳過備案模式。")
        return None, None

    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    form = app.window(handle=hwnd).child_window(auto_id=FORM_AUTO_ID)
    usercode_items = None
    campus = None

    for label, auto_id in TARGETS:
        out("\n" + "=" * 60)
        out(f"{label}  (auto_id={auto_id})  [UIA 展開模式]")
        out("=" * 60)
        try:
            combo = form.child_window(auto_id=auto_id)
            if auto_id == "HOSPITALCODE":
                try:
                    try:
                        sel = combo.iface_value.CurrentValue
                    except Exception:
                        sel = combo.window_text()
                    sel = (sel or "").strip()
                    if sel:
                        campus = sel
                        out(f"目前查詢院區：{sel}")
                except Exception as e:
                    out(f"  讀目前選中的院區失敗: {type(e).__name__}: {e}")
            try:
                combo.expand()
            except Exception:
                pass
            items = [it.window_text() for it in combo.children()]
            items = [t for t in items if t.strip()]
            out(f"共 {len(items)} 個選項：")
            for i, t in enumerate(items):
                out(f"  [{i:3}] {t}")
            if auto_id == "USERCODE":
                usercode_items = items
            try:
                combo.collapse()
            except Exception:
                pass
        except Exception as e:
            out(f"  讀取失敗: {type(e).__name__}: {e}")
    return usercode_items, campus


def main():
    hwnd, title = find_hwnd()
    if not hwnd:
        out("找不到 RIS 主視窗。")
        return
    out(f"主視窗: {title}\n")

    ok, usercode_items, campus = dump_win32(hwnd)
    if not ok:
        out("\n\n!! win32 方式讀不到選項，改用備案。")
        usercode_items, campus = dump_uia_expand(hwnd)

    save_usercode_csv(usercode_items, campus)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f"\n>>> 已存成 {OUT}")


if __name__ == "__main__":
    main()
