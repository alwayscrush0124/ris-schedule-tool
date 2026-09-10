"""
RIS 通用畫面掃描工具
=====================
把目前開著的任何 RIS 畫面的內容倒出來：欄位值 + 表格內容。
這次是要拿「均分池底下有哪些醫師」的資料。

純讀取。不點擊、不輸入、不展開選單、不存檔。

用法：
  1. 在 RIS 裡打開管均分池 / 醫師群組的那個畫面
  2. 讓它顯示出資料（例如點開 W1PLAINW1X光均分，看到底下的成員）
  3. python 5_dump_form.py
  4. 把 form_dump.txt 給我
"""

import io
import re

from pywinauto import Application, Desktop

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
OUT = "form_dump.txt"

_log = io.StringIO()


def out(s=""):
    print(s)
    _log.write(str(s) + "\n")


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
    try:
        kids = [k.window_text() for k in cell.children() if k.window_text()]
        if kids:
            return " ".join(kids)
    except Exception:
        pass
    try:
        return cell.window_text()
    except Exception:
        return ""


def dump_grid(grid, indent="    "):
    """把一個表格的內容印出來。"""
    try:
        rows = grid.children()
    except Exception as e:
        out(f"{indent}(讀取失敗: {e})")
        return

    header, data_rows = None, []
    for r in rows:
        try:
            name = r.window_text()
        except Exception:
            continue
        if "標題" in name or "上方資料列" in name:
            header = r
        elif name.startswith("資料列"):
            data_rows.append(r)

    if header is not None:
        cols = [c.window_text() for c in header.children() if "左上方" not in c.window_text()]
        out(f"{indent}欄位: {' | '.join(cols)}")

    out(f"{indent}共 {len(data_rows)} 列")
    for r in data_rows[:200]:
        vals = []
        for c in r.children():
            t = c.window_text().split()
            if len(t) == 2 and t[0] == "資料列":
                continue
            vals.append(str(cell_value(c)))
        out(f"{indent}  " + " | ".join(vals))
    if len(data_rows) > 200:
        out(f"{indent}  ...(還有 {len(data_rows) - 200} 列未顯示)")


def main():
    hwnd, title = None, None
    for w in Desktop(backend="win32").windows():
        try:
            if re.match(MAIN_TITLE_RE, w.window_text()):
                hwnd, title = w.handle, w.window_text()
                break
        except Exception:
            continue
    if not hwnd:
        out("找不到 RIS 主視窗。")
        return
    out(f"主視窗: {title}\n")

    app = Application(backend="uia").connect(handle=hwnd, timeout=20)
    main_win = app.window(handle=hwnd)

    # 找出所有開著的子畫面。RIS 的畫面代號一律是 f 開頭大寫，例如 fDayWorkSheet。
    print("掃描畫面中，這一步可能要 10~30 秒...")
    all_desc = main_win.descendants(depth=6)
    print(f"  掃到 {len(all_desc)} 個元件")

    forms = []
    for d in all_desc:
        try:
            auto_id = d.element_info.automation_id or ""
        except Exception:
            continue
        if re.match(r"^f[A-Z]", auto_id):
            forms.append((auto_id, d))

    if not forms:
        # 診斷：把所有有代號的元件列出來，看看畫面到底叫什麼
        out("\n!! 沒有比對到 f 開頭的畫面代號。以下是掃到的所有代號，供診斷：")
        seen_ids = {}
        for d in all_desc:
            try:
                aid = d.element_info.automation_id or ""
                ct = str(d.element_info.control_type)
                txt = d.window_text()
            except Exception:
                continue
            if aid and aid not in seen_ids:
                seen_ids[aid] = (ct, txt)
        for aid, (ct, txt) in sorted(seen_ids.items()):
            out(f"  {aid:30} [{ct:14}] {txt[:40]!r}")
        out(f"\n（共 {len(seen_ids)} 個不重複代號）")

    seen = set()
    for auto_id, form in forms:
        if auto_id in seen:
            continue
        seen.add(auto_id)

        out("\n" + "=" * 70)
        out(f"畫面: {form.window_text()}   (auto_id={auto_id})")
        out("=" * 70)

        # 欄位值
        out("\n-- 欄位 --")
        for d in form.descendants():
            try:
                ct = str(d.element_info.control_type)
                aid = d.element_info.automation_id or ""
                txt = d.window_text()
            except Exception:
                continue
            if not aid or aid.startswith("Label"):
                continue
            if any(k in ct for k in ("ComboBox", "TextBox", "DateTimePicker", "CheckBox", "Button")):
                kind = ct.split(".")[-1]
                out(f"  [{kind:16}] {aid:24} = {txt!r}")

        # 表格
        out("\n-- 表格 --")
        n = 0
        for d in form.descendants():
            try:
                ct = str(d.element_info.control_type)
                aid = d.element_info.automation_id or ""
            except Exception:
                continue
            if "DataGridView" in ct:
                n += 1
                out(f"\n  表格 auto_id={aid}")
                dump_grid(d)
        if n == 0:
            out("  （這個畫面沒有表格）")

    if not seen:
        out("找不到任何開著的子畫面。請確認要掃描的畫面是打開的。")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(_log.getvalue())
    print(f"\n>>> 已存成 {OUT}")


if __name__ == "__main__":
    main()
