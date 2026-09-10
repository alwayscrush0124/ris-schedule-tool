"""
RIS 班表掃描工具（唯讀）
========================
用來搞清楚「到底有哪些班別組合、現有對照表接不接得到」，
以及「下拉選單裡有哪些 OFF 帳號」。

*** 這支程式不會修改任何資料。***
它只讀表格。唯一會碰到畫面的是「撈帳號清單」那顆按鈕
（要點一下格子才能展開下拉），撈完會按 ESC 還原並確認格子沒變。

怎麼用：
  1. 在 RIS 查一個條件（某個院區、某段日期），確認畫面有資料
  2. 按「累積這批」
  3. 換個條件再查，再按一次
  4. 累積夠了就看報告，或直接開產生的 CSV

累積的結果會即時寫檔，關掉視窗也不會不見；下次打開會接著累積。

執行： python scan_gui.py
"""

import csv
import os
import queue
import threading
import time
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import ttk, messagebox

from pywinauto.keyboard import send_keys

from ris_gui import (DOCTOR_COL, Ris, load_mapping, norm, plan_for)

COVERAGE_FILE = "掃描_涵蓋率.csv"
ACCOUNTS_FILE = "掃描_帳號候選.csv"

# 組合的鍵：這四欄決定「這是哪一種班」
KEY_COLS = ["檢查地點位址", "統計用分類碼", "排班類別", "檢查室名稱"]

# 看起來像代班/均分帳號的特徵。用來從資料裡認出「已經有人請假」的列，
# 那種列本身就是該院區作業慣例的證據。
OFF_HINTS = ("OFF", "均分", "公用帳號")


def looks_like_off(name):
    up = norm(name).upper()
    return any(h.upper() in up for h in OFF_HINTS)


class Store:
    """累積的掃描結果。鍵是四欄組合，值是次數與看過的 OFF 帳號。"""

    def __init__(self):
        self.rows = {}          # key tuple -> {"n": int, "off": set}
        self.load()

    def load(self):
        if not os.path.exists(COVERAGE_FILE):
            return
        try:
            with open(COVERAGE_FILE, encoding="utf-8-sig") as f:
                for r in csv.DictReader(f):
                    key = tuple(norm(r.get(c, "")) for c in KEY_COLS)
                    off = {x for x in (r.get("曾出現的OFF帳號", "") or "").split("、") if x}
                    self.rows[key] = {"n": int(r.get("出現次數") or 0), "off": off}
        except Exception:
            pass        # 檔案壞掉就當作從頭開始，不要害程式開不起來

    def add(self, row):
        key = tuple(norm(row.get(c, "")) for c in KEY_COLS)
        item = self.rows.setdefault(key, {"n": 0, "off": set()})
        item["n"] += 1
        doc = norm(row.get(DOCTOR_COL, ""))
        if looks_like_off(doc):
            item["off"].add(doc)
        return key

    def save(self, rules):
        with open(COVERAGE_FILE, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(KEY_COLS + ["出現次數", "現有對照表判定", "已涵蓋",
                                   "曾出現的OFF帳號"])
            for key in sorted(self.rows):
                item = self.rows[key]
                row = dict(zip(KEY_COLS, key))
                code, disp, why = plan_for(row, rules)
                w.writerow(list(key) + [item["n"],
                                        (disp or code) if code else why,
                                        "Y" if code else "N",
                                        "、".join(sorted(item["off"]))])

    def stats(self, rules):
        covered = uncovered = 0
        for key in self.rows:
            code, _d, _w = plan_for(dict(zip(KEY_COLS, key)), rules)
            if code:
                covered += 1
            else:
                uncovered += 1
        return covered, uncovered


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RIS 班表掃描（唯讀，不會改任何資料）")
        self.geometry("1150x680")

        self.ris = Ris()
        self.rules = load_mapping()
        self.store = Store()
        self.q = queue.Queue()
        self.busy = False

        self._build()
        self._refresh()
        self.after(100, self._drain)

    def _build(self):
        pad = dict(padx=6, pady=4)

        top = ttk.LabelFrame(self, text="在 RIS 查好一批資料，然後按這裡（可以換條件重複按）")
        top.pack(fill="x", **pad)
        self.btn_scan = ttk.Button(top, text="累積這批", command=self.on_scan)
        self.btn_scan.pack(side="left", padx=6, pady=6)
        self.btn_acct = ttk.Button(top, text="撈帳號清單", command=self.on_accounts)
        self.btn_acct.pack(side="left", padx=6)
        ttk.Label(top, text="（撈帳號會點一下格子展開下拉，撈完自動 ESC 還原）",
                  foreground="#666").pack(side="left", padx=6)
        self.lbl = ttk.Label(top, text="")
        self.lbl.pack(side="left", padx=12)

        mid = ttk.LabelFrame(self, text="累積到目前為止的班別組合")
        mid.pack(fill="both", expand=True, **pad)
        cols = KEY_COLS + ["出現次數", "現有對照表判定", "曾出現的OFF帳號"]
        self.tree = ttk.Treeview(mid, columns=cols, show="headings")
        widths = {"檢查地點位址": 90, "統計用分類碼": 90, "排班類別": 130,
                  "檢查室名稱": 150, "出現次數": 70,
                  "現有對照表判定": 200, "曾出現的OFF帳號": 200}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths.get(c, 100), anchor="w")
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("miss", foreground="#b00")

        logf = ttk.LabelFrame(self, text="訊息")
        logf.pack(fill="x", **pad)
        self.txt = tk.Text(logf, height=8, wrap="word")
        self.txt.pack(fill="both", expand=True)
        self.log(f"對照表載入 {len(self.rules)} 條規則。"
                 f"已累積 {len(self.store.rows)} 種組合。")
        self.log("這支程式只讀資料，不會修改也不會存檔。")

    def log(self, s):
        self.txt.insert("end", f"{datetime.now():%H:%M:%S}  {s}\n")
        self.txt.see("end")

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for key in sorted(self.store.rows):
            item = self.store.rows[key]
            code, disp, why = plan_for(dict(zip(KEY_COLS, key)), self.rules)
            self.tree.insert("", "end",
                             tags=() if code else ("miss",),
                             values=list(key) + [item["n"],
                                                 (disp or code) if code else why,
                                                 "、".join(sorted(item["off"]))])
        covered, uncovered = self.store.stats(self.rules)
        self.lbl.config(text=f"共 {len(self.store.rows)} 種組合："
                             f"已涵蓋 {covered}、缺 {uncovered}")

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "refresh":
                    self._refresh()
                elif kind == "done":
                    self.busy = False
                    self.btn_scan.config(state="normal")
                    self.btn_acct.config(state="normal")
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _run(self, fn):
        if self.busy:
            return
        self.busy = True
        self.btn_scan.config(state="disabled")
        self.btn_acct.config(state="disabled")

        def wrap():
            try:
                fn()
            except Exception as e:
                self.q.put(("log", f"!! {type(e).__name__}: {e}"))
                self.q.put(("log", traceback.format_exc()))
            finally:
                self.q.put(("done", None))

        threading.Thread(target=wrap, daemon=True).start()

    # ---------- 累積 ----------
    def on_scan(self):
        def job():
            self.q.put(("log", "讀取畫面表格中..."))
            self.ris.connect()
            t = time.time()
            rows = self.ris.read_rows()
            if not rows:
                self.q.put(("log", "表格是空的 —— 請先在 RIS 查詢出資料。"))
                return

            before = len(self.store.rows)
            new_keys = []
            for r in rows:
                key = self.store.add(r)
                if len(self.store.rows) > before + len(new_keys):
                    new_keys.append(key)

            self.store.save(self.rules)
            self.q.put(("refresh", None))
            self.q.put(("log", f"讀到 {len(rows)} 列（{time.time() - t:.1f} 秒），"
                               f"新增 {len(self.store.rows) - before} 種組合。"))
            for key in new_keys:
                code, disp, why = plan_for(dict(zip(KEY_COLS, key)), self.rules)
                mark = f"-> {disp or code}" if code else f"** 缺規則：{why}"
                self.q.put(("log", f"    新組合 {' / '.join(key)}  {mark}"))
            self.q.put(("log", f"已存到 {COVERAGE_FILE}"))
        self._run(job)

    # ---------- 撈帳號 ----------
    def on_accounts(self):
        if not messagebox.askyesno(
                "撈帳號清單",
                "會點一下表格裡「排班醫師」的第一格、展開下拉選單、"
                "把項目全部讀下來，然後按 ESC 還原。\n\n"
                "不會輸入任何東西，也不會存檔。\n"
                "撈完會確認那一格的值沒有變。\n\n要繼續嗎？"):
            return

        def job():
            self.ris.connect()
            rows = self.ris.read_rows()
            if not rows:
                self.q.put(("log", "表格是空的，請先查詢出資料（要有列才點得到格子）。"))
                return

            idx = rows[0]["_row"]
            before = self.ris.cell_text(idx)
            self.q.put(("log", f"用第 {idx} 列展開下拉（目前值 {before!r}）..."))

            ok, how = self.ris.goto_row(idx)
            if not ok:
                self.q.put(("log", f"點不進格子（{how}），放棄。沒有動到任何東西。"))
                return

            send_keys("{F4}")
            time.sleep(1.0)
            items = []
            try:
                for it in self.ris.win.descendants(control_type="ListItem"):
                    name = norm(it.element_info.name)
                    if name:
                        items.append(name)
            except Exception as e:
                self.q.put(("log", f"讀取清單失敗：{type(e).__name__}"))

            send_keys("{ESC}")
            time.sleep(0.3)
            send_keys("{ESC}")
            time.sleep(0.4)

            after = self.ris.cell_text(idx)
            if after != before:
                self.q.put(("log", f"!! 那一格從 {before!r} 變成 {after!r}。"))
                self.q.put(("log", "!! 請回 RIS 重新查詢，提示選「否」還原。"))
                return
            self.q.put(("log", f"格子沒有被改到（還是 {after!r}）。"))

            seen, uniq = set(), []
            for name in items:
                if name not in seen:
                    seen.add(name)
                    uniq.append(name)

            with open(ACCOUNTS_FILE, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["顯示名稱", "推測帳號", "像不像代班帳號", "備註"])
                for name in uniq:
                    guess, note = "", ""
                    if name.endswith("均分"):
                        # CCTAOFF均分 -> CCTAOFF。但 W1PLAINW1X光均分 這種
                        # 是「帳號+姓名」黏在一起，去掉「均分」不等於帳號，
                        # 所以一律標成推測，不要當成定論。
                        guess = name[:-2]
                        note = "推測：去掉「均分」；請自己確認"
                    elif looks_like_off(name):
                        note = "看起來是代班/公用帳號，但拆不出帳號，要人工判斷"
                    w.writerow([name, guess, "Y" if looks_like_off(name) else "",
                                note])

            off_n = sum(1 for n in uniq if looks_like_off(n))
            self.q.put(("log", f"撈到 {len(uniq)} 個項目，其中 {off_n} 個像代班帳號。"))
            self.q.put(("log", f"已存到 {ACCOUNTS_FILE}"))
        self._run(job)


if __name__ == "__main__":
    App().mainloop()
