"""
RIS 班表代班小工具（GUI）
==========================
半自動：條件你選，程式只負責把重複的修改跑完。

流程：
  1. 在 RIS 的「醫師班表維護」查詢出資料
  2. 這裡按「讀取畫面表格」
  3. 用上面的篩選條件縮小範圍（醫師 / 日期 / 分類 / 類別 / 時段）
  4. 選「改成」哪個帳號
  5. 按「開始修改」
  6. 回 RIS 檢查，自己按存檔

*** 這個程式永遠不會替你按存檔。***

執行： python ris_gui.py
"""

import csv
import ctypes
import os
import queue
import re
import sys
import threading
import time
import traceback
import tkinter as tk
from datetime import datetime
from tkinter import ttk, messagebox

from pywinauto import Application, Desktop, mouse
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys
from pywinauto.uia_defines import IUIA
from pywinauto.uia_element_info import UIAElementInfo

# 2026-08-14：路徑基準要分兩種，不能只用一個 —— 這是實測（用一支小
# 探測腳本 path_probe.py 分別編 --onefile 量出來）才確定的，不是猜的：
#
#   OUTPUT_BASE_DIR（sys.argv[0] 所在資料夾）：給「要留下來給人看」的
#   輸出檔用（log、本機設定）。
#   DATA_BASE_DIR（__file__ 所在資料夾，找不到才退回 OUTPUT_BASE_DIR）：
#   給「隨程式一起帶著走」的唯讀設定檔用（設定/、accounts.csv...）。
#
# 三種啟動方式下，這兩個基準分別是：
#   1. `python ris_gui.py`：兩個都等於這支 .py 的所在資料夾，沒差別。
#   2. Nuitka --standalone：兩個都等於 ris_gui.exe 在 dist 資料夾裡的位置，
#      沒差別（standalone 不會把自己解壓縮到別的地方）。
#   3. Nuitka --onefile：**兩個不一樣**——exe 執行時會把自己解壓縮到一個
#      執行完就會被刪掉的暫存資料夾再啟動「真正的」子程式：
#        sys.argv[0] 被刻意設成使用者雙擊的那個 exe 路徑（原本的位置），
#          隨附的設定檔並不在那裡；但正因為它「原本的位置」不會消失，
#          拿來放輸出檔（log、本機設定）才對，程式關掉還留得住。
#        __file__ 指向那個暫存解壓縮資料夾，隨附的設定檔真的在那裡，
#          但程式一關那個資料夾就沒了，拿來放輸出檔會憑空消失。
#      2026-08-14 實測踩到：全部混用 sys.argv[0] 那組，開起來院區選單和
#      醫師清單整個是空的，因為找的是使用者雙擊 exe 的原始位置，根本沒有
#      設定檔在那裡（設定檔其實在暫存資料夾）。
OUTPUT_BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
try:
    DATA_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    DATA_BASE_DIR = OUTPUT_BASE_DIR
if not os.path.isdir(os.path.join(DATA_BASE_DIR, "設定")):
    # __file__ 那邊沒有設定檔（例如日後 Nuitka 行為變了），退回另一組，
    # 至少不會比原本的裸相對路徑更糟。
    DATA_BASE_DIR = OUTPUT_BASE_DIR

MAIN_TITLE_RE = r".*放射線資訊管理系統.*"
FORM_AUTO_ID = "fDayWorkSheet"
GRID_AUTO_ID = "D"
DOCTOR_COL = "排班醫師"
ACCOUNTS_FILE = "accounts.csv"
DOCTORS_FILE = "doctors.csv"
OPTIONS_FILE = "欄位選項.csv"
MAPPING_FILE = "mapping.csv"
LOG_FILE = os.path.join(OUTPUT_BASE_DIR, "gui_log.txt")

# 各院區的設定放在 設定/<院區>/ 底下。
# 找不到這個資料夾就退回讀根目錄的 accounts.csv / mapping.csv（舊的用法）。
CONFIG_DIR = os.path.join(DATA_BASE_DIR, "設定")
LOCAL_SETTINGS = os.path.join(OUTPUT_BASE_DIR, "本機設定.txt")  # 記住上次選的院區

COLS = ["起始日期", "起始時間", "結束日期", "結束時間", "排班時段",
        "檢查室名稱", "門急住限制", "統計用分類碼", "排班類別",
        "檢查地點位址", "排班醫師"]

# 表格上顯示哪幾欄（太多欄會看不完）
SHOW = ["起始日期", "排班時段", "統計用分類碼", "排班類別",
        "檢查室名稱", "門急住限制", "檢查地點位址", "排班醫師"]

# mapping.csv 比對用的三欄
MATCH_COLS = ("統計用分類碼", "排班類別", "檢查室名稱")

# 等待上限（秒）。這些是「最多等這麼久」，不是固定睡這麼久 ——
# 條件一成立就馬上往下走，所以機器快的時候不會變慢。
# RIS 忙起來的時候一列要 8 秒以上，設太小會把成功的寫入誤判成失敗。
AUTOCOMPLETE_WAIT = 4.0     # 打完字等下拉補完
COMMIT_WAIT = 0.6           # 每按一次 Enter 之後等格子反映新值。
                            # 第一次註定不會過（Enter 被吃掉），所以這個值
                            # 每列都會被付滿一次，不要設太大。
COMMIT_ENTER_TRIES = 3      # Enter 最多按幾次（RIS 會吃掉第一個，見 write_cell）
MAX_SCROLL_STEPS = 15       # 表格離目前捲動位置太遠時，最多捲幾格滾輪去找那一列

# 可以拿來篩選的欄位
FILTERS = ["排班醫師", "起始日期", "統計用分類碼", "排班類別", "排班時段", "門急住限制"]

# 允許修改的欄位。
# 日期與時間沒放進來 —— 那等於把班移到別天，風險跟性質都不一樣，
# 要開放的話應該另外設計，不是混在這裡。
EDITABLE = [DOCTOR_COL, "排班時段", "統計用分類碼", "排班類別",
            "檢查室名稱", "門急住限制", "檢查地點位址"]


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


# send_keys 會把這些字元當成語法（分組、修飾鍵、重複次數…），
# 要當一般文字送就得包成 {x}。
SEND_KEYS_SPECIAL = "^+%~(){}[]"


WM_CHAR = 0x0102


def type_text(text):
    """用 send_keys 打字。只適合純 ASCII。

    兩個坑都在 send_keys 的預設值裡：
      * ( ) ^ + % ~ { } [ ] 是語法字元，不跳脫的話會被吃掉。
      * with_spaces 預設 False，字串裡的空白會被直接忽略。
    排班醫師的帳號剛好都是英數字（CCTAOFF），所以一直沒踩到；
    改到「急診夜間值班A (21:30-07:59)」這種值才爆出來（2026-07-23 實測）。
    """
    escaped = "".join("{" + c + "}" if c in SEND_KEYS_SPECIAL else c
                      for c in text)
    send_keys(escaped, with_spaces=True)


def post_chars(text):
    """把字元用 WM_CHAR 直接送給目前焦點的原生視窗。成功回 True。

    *** 中文一定要走這條。***
    send_keys 送的是 SendInput 的 Unicode 封包，這個 WinForms 控制項不吃 ——
    實測（2026-07-23）打「上半天」毫無反應，打 'B' 立刻命中 BOD01，
    展開下拉再打、先按 Ctrl+A 再打，通通沒用。
    改用 WM_CHAR 直接送字元訊息就成功了。

    對控制項而言 WM_CHAR 跟真的敲鍵盤是同一種輸入，
    所以它跟 send_keys 一樣會讓 RIS 認定這一格被編輯過
    （這點跟「直接設值」有本質差別，別混為一談）。
    """
    try:
        hwnd = IUIA().iuia.GetFocusedElement().CurrentNativeWindowHandle
    except Exception:
        hwnd = 0
    if not hwnd:
        return False
    user32 = ctypes.windll.user32
    for ch in text:
        user32.PostMessageW(hwnd, WM_CHAR, ord(ch), 0)
        time.sleep(0.01)
    return True


def type_into_editor(text):
    """把值打進目前開著的編輯器。ASCII 走 send_keys，其他走 WM_CHAR。

    ASCII 那條是今天實際存檔驗證過的路徑，所以能用就用；
    非 ASCII 只有 WM_CHAR 送得進去。
    """
    if all(ord(c) < 128 for c in text):
        type_text(text)
        return
    if not post_chars(text):
        type_text(text)        # 拿不到原生 handle 就死馬當活馬醫


# ----------------------------------------------------------------------
# 跟 RIS 溝通的部分
# ----------------------------------------------------------------------

class Ris:
    def __init__(self):
        self.grid = None
        # 讀表格時一次建好的格子索引 {列號: {欄名: 格子元件}}。
        # 每一次跟 RIS 要元件都是跨程序呼叫，很貴 ——
        # 以前每改一列要重新列舉整個表格 3~4 次，一列就耗掉好幾秒。
        self.cells = {}
        # 列元件 {列號: 元件}。重抓某一格時只要問這一列的子元件（約 0.1 秒），
        # 不必重新列舉整張表格（實測 3.9 秒）。
        self.row_elems = {}
        # 上一次寫入各階段花的時間 (打字, 等補完, 提交, 按了幾次Enter)，
        # 寫進 log 用。慢的時候才知道要往哪裡查。
        self.last_timing = (0, 0, 0, 0)
        # 2026-08-10 效能診斷用，更細的拆解。只加欄位、不改既有邏輯，
        # 診斷完不需要了可以整段刪掉，不影響其他部分。
        # goto_row: (找格子, 進迴圈前的編輯器檢查, 捲動, 對焦, 點擊,
        #            等編輯器開, 第幾次點擊才成功)
        self.last_goto_detail = (0, 0, 0, 0, 0, 0, 0)
        # 等編輯器開那段迴圈裡：(問了幾次, 問的總耗時, 單次問最久)
        self.last_editor_poll_detail = (0, 0, 0)
        # write_cell 提交段: (第1個Enter, 第2個Enter, 第一輪等格子反映)
        self.last_commit_detail = (0, 0, 0)
        # 等格子反映新值那段迴圈裡：(問了幾次, 問的總耗時, 單次問最久)
        self.last_wait_value_detail = (0, 0, 0)

    def connect(self):
        hwnd = None
        for w in Desktop(backend="win32").windows():
            try:
                if re.match(MAIN_TITLE_RE, w.window_text()):
                    hwnd = w.handle
                    break
            except Exception:
                continue
        if not hwnd:
            raise RuntimeError("找不到 RIS 主視窗，請確認 RIS 開著。")

        app = Application(backend="uia").connect(handle=hwnd, timeout=20)
        self.win = app.window(handle=hwnd)
        self.grid = (self.win.child_window(auto_id=FORM_AUTO_ID)
                     .child_window(auto_id=GRID_AUTO_ID))
        self.grid.wait("exists", timeout=15)

    @staticmethod
    def _cell_value(cell):
        """讀一格的值。

        legacy_properties() 一次呼叫就能同時拿到值和名稱，
        比先問 ValuePattern 再問名稱少跑兩趟，量大時差很多。
        """
        try:
            v = cell.legacy_properties().get("Value")
            if v not in (None, ""):
                return v
        except Exception:
            pass
        try:
            v = cell.iface_value.CurrentValue
            if v not in (None, ""):
                return v
        except Exception:
            pass
        return cell.window_text()

    @staticmethod
    def _col_of(cell):
        p = cell.window_text().split()
        return p[0] if p else ""

    def read_rows(self, progress=None):
        """把表格內容讀成一串 dict。

        速度考量：每一次跟 RIS 要資料都是跨程序呼叫，很貴。
        所以這裡不再逐格去問「你是哪一欄」，而是靠位置判斷 ——
        每一列的第一格固定是列首，後面依序就是 COLS 的順序。
        只在第一列驗證一次順序，確認沒變才繼續。
        """
        rows = []
        self.cells = {}
        self.row_elems = {}
        data_rows = []
        for c in self.grid.children():
            name = c.window_text()
            if name.startswith("資料列"):
                data_rows.append((int(name.split()[1]), c))

        total = len(data_rows)
        checked = False

        for i, (idx, r) in enumerate(data_rows):
            kids = r.children()
            if len(kids) < len(COLS) + 1:
                continue

            if not checked:
                # 只做一次：確認第一格真的是列首、欄位順序跟預期一致
                got = [k.window_text().split()[0] for k in kids[:len(COLS) + 1]]
                if got[0] != "資料列" or got[1:] != COLS:
                    raise RuntimeError(
                        "表格欄位順序跟預期不符，為安全起見中止。\n"
                        f"讀到的順序：{got}")
                checked = True

            cols = kids[1:len(COLS) + 1]
            vals = [self._cell_value(k) for k in cols]
            row = dict(zip(COLS, vals))
            # 表格最後那列是空白的「新增列」，不能讓它混進清單 ——
            # 萬一被選到，程式會去點它、往裡面打字，等於生出一筆新記錄。
            # 不能只看「所有欄位都空」，它有時候會有一兩欄不是空的，
            # 所以改看幾個關鍵欄位：沒有日期也沒有醫師的就不是真資料。
            blank = ("", "(null)")
            if all(norm(v) in blank for v in row.values()):
                continue
            if norm(row["起始日期"]) in blank and norm(row[DOCTOR_COL]) in blank:
                continue
            row["_row"] = idx
            rows.append(row)
            self.cells[idx] = dict(zip(COLS, cols))
            self.row_elems[idx] = r
            if progress and (i + 1) % 5 == 0:
                progress(i + 1, total)
        return rows

    def _find_cell(self, idx, col):
        """拿某一格。優先用讀表格時建好的索引，找不到才回頭掃一次。

        索引裡的元件可能過期（使用者在 RIS 重新查詢過）。
        過期的元件會丟 COMError，接住之後改用掃的；
        真的變了的話，寫入前的內容核對會擋下來。
        """
        cell = self.cells.get(idx, {}).get(col)
        if cell is not None:
            try:
                cell.window_text()      # 便宜的存活測試
                return cell
            except Exception:
                self.cells.pop(idx, None)

        # 快取沒了，先問那一列的子元件（約 0.1 秒），
        # 真的不行才退回列舉整張表格（3.9 秒）。
        row = self.row_elems.get(idx)
        if row is not None:
            try:
                for c in row.children():
                    if self._col_of(c) == col:
                        self.cells.setdefault(idx, {})[col] = c
                        return c
            except Exception:
                self.row_elems.pop(idx, None)

        for r in self.grid.children():
            name = r.window_text()
            if name.startswith("資料列") and name.split()[1:] == [str(idx)]:
                for c in r.children():
                    if self._col_of(c) == col:
                        return c
        return None

    def read_row(self, idx):
        """重新讀某一列，用來在修改前核對。"""
        cached = self.cells.get(idx)
        if cached:
            try:
                return {c: self._cell_value(cached[c]) for c in COLS}
            except Exception:
                self.cells.pop(idx, None)

        for r in self.grid.children():
            name = r.window_text()
            if name.startswith("資料列") and name.split()[1:] == [str(idx)]:
                kids = r.children()
                if len(kids) < len(COLS) + 1:
                    return None
                vals = [self._cell_value(k) for k in kids[1:len(COLS) + 1]]
                return dict(zip(COLS, vals))
        return None

    def _editor_on(self, cell):
        """這一格的「編輯器」開著而且焦點在上面嗎？

        要問的是編輯器，不是「這一格是不是目前格」——
        兩者不一樣，而且差別會咬人：
        按下 Enter 之後游標會自動落到下一列的同一欄，那一格變成目前格，
        但**編輯器沒有開**。這時候直接打字會打進空氣裡。
        （2026-07-23 實測：批次最後一列因此失敗，補完讀回來是空字串。）

        不能問「焦點元件叫什麼名字」—— 編輯器叫「編輯控制項」，
        掛在「編輯面板」底下，往上追完全看不到列號。（也實測過。）

        所以判斷方式是：焦點元件要有 ValuePattern（＝真的是編輯器），
        而且它的位置正好蓋在這一格上。
        """
        try:
            el = IUIA().iuia.GetFocusedElement()
            UIAWrapper(UIAElementInfo(el)).iface_value   # 不是編輯器就會丟例外
            fr = el.CurrentBoundingRectangle
            cr = cell.rectangle()
            return (abs((fr.top + fr.bottom) // 2 - (cr.top + cr.bottom) // 2) <= 5
                    and abs(fr.left - cr.left) <= 5)
        except Exception:
            return False

    def goto_row(self, idx, col=DOCTOR_COL):
        """點進第 idx 列的某一格，並確認編輯器真的開在那一格上。

        一律用點擊，不用方向鍵跨列移動 —— 因為格子的編輯器是
        「下拉式方塊」，方向鍵經過中途的列時有可能改到那些列的值，
        而那些列不在我們的驗證名單裡，改壞了不會被發現。

        點擊實測會偶發沒生效（點了但編輯器還停在原本那一列），
        所以要驗證 + 重試，不到位就回報失敗讓上層中止。

        回傳 (成功與否, 說明)
        """
        # 2026-08-10 效能診斷：每次呼叫先歸零，避免殘留上一列的數字。
        self.last_goto_detail = (0, 0, 0, 0, 0, 0, 0)
        self.last_editor_poll_detail = (0, 0, 0)

        t0 = time.time()
        cell = self._find_cell(idx, col)
        t_find = time.time() - t0
        if cell is None:
            return False, f"找不到第 {idx} 列的「{col}」格"
        t0 = time.time()
        early_on = self._editor_on(cell)
        t_early_on = time.time() - t0
        if early_on:
            self.last_goto_detail = (t_find, t_early_on, 0, 0, 0, 0, 0)
            return True, "編輯器已經開著"

        for attempt in range(3):
            t0 = time.time()
            try:
                cell.iface_scrollitem.ScrollIntoView()
            except Exception:
                pass
            t_scroll = time.time() - t0

            t0 = time.time()
            try:
                self.win.set_focus()
            except Exception as e:
                return False, f"點選失敗: {type(e).__name__}"
            t_focus = time.time() - t0

            t0 = time.time()
            try:
                cell.click_input()
            except Exception as e:
                return False, f"點選失敗: {type(e).__name__}"
            t_click = time.time() - t0

            t0 = time.time()
            deadline = time.time() + 1.0
            n_polls, t_polls, t_poll_max = 0, 0.0, 0.0
            while time.time() < deadline:
                tp0 = time.time()
                on = self._editor_on(cell)
                dtp = time.time() - tp0
                n_polls += 1
                t_polls += dtp
                t_poll_max = max(t_poll_max, dtp)
                if on:
                    self.last_goto_detail = (t_find, t_early_on, t_scroll, t_focus,
                                              t_click, time.time() - t0, attempt + 1)
                    self.last_editor_poll_detail = (n_polls, t_polls, t_poll_max)
                    return True, f"點擊{attempt + 1}次"
                time.sleep(0.03)
            self.last_goto_detail = (t_find, t_early_on, t_scroll, t_focus,
                                      t_click, time.time() - t0, attempt + 1)
            self.last_editor_poll_detail = (n_polls, t_polls, t_poll_max)

        # 表格是虛擬化的：離目前捲動位置太遠的列，ScrollIntoView 叫不動、
        # 點了也沒用，因為那一列根本沒被畫出來（不是格子找不到，是格子
        # 在畫面外）。2026-07-30 用台大總院 135 列的表實測踩到：第 0/10/30
        # 列都正常，隔比較遠的第 60/90/120 列點 3 次都開不了編輯器。
        #
        # 先試垂直捲軸的「向下翻頁／向上翻頁」按鈕 —— 實測一次翻頁的幅度
        # 遠比滑鼠滾輪一格大很多（從第 60 列附近翻一頁就到第 90 列了），
        # 長距離跳頁效率好很多。兩個方向都試，因為不確定目標列在上面
        # 還是下面，也可能翻頁翻過頭。
        # 找不到翻頁按鈕、或翻頁翻不到（例如剛好卡在按鈕翻不動的邊界）
        # 才退回滑鼠中鍵滾輪一格一格捲，當最後手段做微調。
        # 這整段都只有捲動和點擊，不會打字，安全，頂多是點不到。
        for how_scroll, step_fn in self._scroll_strategies():
            for direction in (-1, 1):
                for _step in range(MAX_SCROLL_STEPS):
                    if not step_fn(direction):
                        break
                    time.sleep(0.25)

                    cell = self._find_cell(idx, col)
                    if cell is None:
                        continue
                    try:
                        cell.iface_scrollitem.ScrollIntoView()
                    except Exception:
                        pass
                    try:
                        self.win.set_focus()
                        cell.click_input()
                    except Exception:
                        continue

                    deadline = time.time() + 1.0
                    while time.time() < deadline:
                        if self._editor_on(cell):
                            way = "下" if direction < 0 else "上"
                            return True, f"{how_scroll}{way}{_step + 1}次後點到"
                        time.sleep(0.03)

        return False, (f"點了 3 次、翻頁和滾輪兩個方向都試過，"
                        f"第 {idx} 列「{col}」的編輯器還是沒開起來")

    def _scroll_strategies(self):
        """回傳 [(說明, 捲動函式), ...]，捲動函式(direction) -> 有沒有真的捲到。

        先翻頁（幅度大，效率好），翻頁按鈕找不到才退回滾輪（幅度小，
        當微調或最後手段）。direction: -1 表示往下/往後，1 表示往上/往前。
        """
        strategies = []
        try:
            bar = self.grid.descendants(control_type="ScrollBar")[0]
            pgdn = bar.children(title="向下翻頁")[0]
            pgup = bar.children(title="向上翻頁")[0]

            def page(direction):
                btn = pgdn if direction < 0 else pgup
                try:
                    btn.invoke()
                except Exception:
                    try:
                        btn.click_input()
                    except Exception:
                        return False
                return True

            strategies.append(("翻頁", page))
        except Exception:
            pass

        def wheel(direction):
            try:
                grid_rect = self.grid.rectangle()
                cx = (grid_rect.left + grid_rect.right) // 2
                cy = (grid_rect.top + grid_rect.bottom) // 2
                mouse.scroll(coords=(cx, cy), wheel_dist=direction)
                return True
            except Exception:
                return False

        strategies.append(("滾輪", wheel))
        return strategies

    def cell_text(self, idx, col=DOCTOR_COL):
        cell = self._find_cell(idx, col)
        return norm(self._cell_value(cell)) if cell is not None else "(讀不到)"

    def refresh_cell(self, idx, col=DOCTOR_COL):
        """丟掉快取的格子元件，重新抓一個回來。

        提交之後 RIS 有時候會換掉整列的元件物件。舊的那個還活著
        （存活測試過得了）但已經跟畫面脫鉤，問它幾次都回答舊值 ——
        看起來很像「系統很慢」，其實再等也不會變。
        實測（2026-07-23）輪詢舊元件 3 秒都是舊值，同一時間重讀整張表卻是新值。
        """
        row = self.cells.get(idx)
        if row:
            row.pop(col, None)          # 快取沒了，_find_cell 就會重新掃描
        cell = self._find_cell(idx, col)
        if cell is not None:
            self.cells.setdefault(idx, {})[col] = cell
        return cell

    @staticmethod
    def set_editor_text(text):
        """直接把值塞進編輯器（不是塞進格子）。成功回 True。

        *** 這跟「對格子 SetValue」是兩回事，不要搞混。***
        對格子 SetValue：RIS 不認帳，存檔時整列被跳過（早上實測踩過）。
        對編輯器 SetValue：格子已經在編輯狀態，設的是編輯中的內容，
                          再按 Enter 提交 —— 跟人打字進去的路徑一樣。

        用途是打字送不進去的值。send_keys 送不了中文（實測：打「上半天」
        毫無反應，打 'B' 立刻命中 BOD01），而班表上很多值是中文。
        """
        try:
            el = IUIA().iuia.GetFocusedElement()
            UIAWrapper(UIAElementInfo(el)).iface_value.SetValue(text)
            return True
        except Exception:
            return False

    @staticmethod
    def editor_text():
        """格子進入編輯狀態後，那個下拉式方塊目前顯示的字。"""
        try:
            el = IUIA().iuia.GetFocusedElement()
            return norm(UIAWrapper(UIAElementInfo(el)).iface_value.CurrentValue)
        except Exception:
            return ""

    def write_cell(self, idx, col, code, display=None):
        """像真人一樣把值打進某一格。回傳 (成功, 看到的值, 失敗原因)。

        col 預設是排班醫師，但其他欄位（排班時段、統計用分類碼、排班類別、
        檢查室名稱、門急住限制、檢查地點位址）走的是同一套 ——
        它們的編輯器都是同一種下拉式方塊。

        *** 不可以改用 ValuePattern.SetValue。***
        用 SetValue 設的值 RIS 不認帳：畫面會變、讀回來也對，
        但存檔時整列被跳過，重新查詢就跳回舊值。
        實測（2026-07-23 正式環境）一次改 9 列，只有 1 列進得去。
        只有「從鍵盤打進去」才會讓 RIS 認定這一列被編輯過。

        流程：全選 -> 打帳號 -> 下拉自動補完成顯示名稱 -> 檢查 -> Enter 提交。
        補完的結果不如預期就按 ESC 取消，格子回到原值，絕不提交。

        呼叫前游標要已經在這一格（goto_row 負責）。
        """
        # 2026-08-10 效能診斷：每次呼叫先歸零，避免殘留上一列的數字。
        self.last_commit_detail = (0, 0, 0)
        before = self.cell_text(idx, col)
        # 「已經是目標值」的判斷要跟後面補完比對用同一套嚴不嚴格規則。
        # 之前這裡永遠用 startswith，害「LDCT國健」被誤判成已經是「LDCT」
        # （前者恰好以後者開頭）而直接跳過，畫面看起來沒動、其實根本沒改到
        # ——2026-07-30 用排班類別 LDCT/LDCT國健、MG/MG國健、
        # Pediatrics/Pediatrics ortho 這幾組同名碰撞實測踩到。
        want0 = norm(display) if display else norm(code)
        already = (before == want0) if display else before.startswith(code)
        if already:
            return True, before, ""          # 已經是目標值了

        t0 = time.time()
        send_keys("^a")
        type_into_editor(code)
        t_type = time.time() - t0
        t0 = time.time()

        # 知道期待顯示什麼的時候就要求完全相符，不能只看開頭 ——
        # 下拉是前綴比對，會有碰撞：實測打 'CHE'（要 CHE胸部放射(公用)）
        # 補成了 'Chest CT手動均分'。同理打 'X1' 可能補成 'X10均分班10'，
        # 那就會把班改給錯的均分池，而且開頭剛好吻合、檢查不出來。
        want = norm(display) if display else norm(code)
        strict = bool(display)

        def matched(text):
            return text == want if strict else text.startswith(code)

        deadline = time.time() + AUTOCOMPLETE_WAIT
        shown = ""
        while time.time() < deadline:
            shown = self.editor_text()
            if matched(shown):
                break
            time.sleep(0.05)
        t_auto = time.time() - t0

        # 打字沒送到 -> 改用直接設編輯器的值。
        # send_keys 送不了中文，而班表上很多值是中文（實測 2026-07-23）。
        # 打字是已經驗證過能存進資料庫的路徑，所以先打字、失敗才退這一步。
        if not matched(shown):
            if self.set_editor_text(want):
                deadline = time.time() + AUTOCOMPLETE_WAIT
                while time.time() < deadline:
                    shown = self.editor_text()
                    if matched(shown):
                        break
                    time.sleep(0.05)
            t_auto = time.time() - t0

        # 兩條路都放不進去（打錯值、選單裡沒這一項…）-> 取消，不留痕跡
        if not matched(shown):
            send_keys("{ESC}")
            time.sleep(0.3)
            if not shown:
                return False, "", "編輯器沒開起來，字沒打進去（已 ESC，這一列沒被改到）"
            return False, shown, (f"下拉補成 {shown!r}，預期 {want!r}"
                                  "（已 ESC，這一列沒被改到）")

        # 提交。重點是「多按幾次 Enter，每次都檢查」，不是「按一次然後等很久」。
        #
        # RIS 會吃掉第一個 Enter —— 多半是自動補完的下拉還開著，
        # 那個 Enter 只用來關掉下拉，沒有提交到格子。
        # 實測（2026-07-23）按一次 Enter 然後等 3 秒、6 秒都一樣失敗，
        # 因為根本沒送出去，等再久也不會變。
        #
        # 「讀到舊值」還有第二種原因：快取的元件跟畫面脫鉤了，
        # 問它幾次都回答舊值。所以第一次沒過時順便換一個元件再問。
        #
        # 多按的 Enter 是安全的：Enter 在非編輯狀態只會把游標往下移一列，
        # 不會像方向鍵那樣改到下拉選單的值。
        # 第一個 Enter 被吃掉是常態不是異常（多半只用來關掉自動補完的下拉），
        # 所以直接連按兩下再檢查，省掉一整輪註定失敗的等待。
        # 多按的 Enter 是安全的：非編輯狀態下 Enter 只會把游標往下移一列，
        # 不會像方向鍵那樣改到下拉選單的值。
        #
        # 「重抓元件」這種比較貴的手段留到 Enter 都按完還是不對時才用。
        t0 = time.time()
        now = ""
        t_e0 = time.time()
        send_keys("{ENTER}")
        t_enter1 = time.time() - t_e0
        t_e0 = time.time()
        send_keys("{ENTER}")
        t_enter2 = time.time() - t_e0
        for i in range(COMMIT_ENTER_TRIES - 1):
            t_w0 = time.time()
            ok, now = self._wait_value(idx, col, code, want, COMMIT_WAIT, strict)
            if i == 0:
                # 2026-08-10 效能診斷：只記第一輪，那一輪是每列都會發生的。
                self.last_commit_detail = (t_enter1, t_enter2, time.time() - t_w0)
            if ok:
                self.last_timing = (t_type, t_auto, time.time() - t0, i + 2)
                return True, now, ""
            send_keys("{ENTER}")

        self.refresh_cell(idx, col)
        ok, now = self._wait_value(idx, col, code, want, COMMIT_WAIT, strict)
        self.last_timing = (t_type, t_auto, time.time() - t0, COMMIT_ENTER_TRIES + 1)
        if ok:
            return True, now, ""

        return False, now, (f"按了 {COMMIT_ENTER_TRIES + 1} 次 Enter（含重抓元件），"
                            f"格子還是 {now!r}")

    def _wait_value(self, idx, col, code, want, timeout, strict=True):
        """等這一格的值變成目標。回傳 (成功, 最後讀到的值)。"""
        deadline = time.time() + timeout
        now = ""
        # 2026-08-10 效能診斷：問了幾次、每次問 cell_text() 花多久。
        n_polls, t_polls, t_poll_max = 0, 0.0, 0.0
        while True:
            tp0 = time.time()
            now = self.cell_text(idx, col)
            dtp = time.time() - tp0
            n_polls += 1
            t_polls += dtp
            t_poll_max = max(t_poll_max, dtp)
            if (now == want) if strict else now.startswith(code):
                self.last_wait_value_detail = (n_polls, t_polls, t_poll_max)
                return True, now
            if time.time() >= deadline:
                self.last_wait_value_detail = (n_polls, t_polls, t_poll_max)
                return False, now
            time.sleep(0.05)


def load_accounts(path=ACCOUNTS_FILE):
    """讀帳號清單。回傳 [(帳號, 顯示名稱, 已驗證, 說明), ...]"""
    items = []
    if not os.path.exists(path):
        return items
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            p = [x.strip() for x in next(csv.reader([line]))]
            if p[0] == "帳號":
                continue
            while len(p) < 4:
                p.append("")
            items.append(tuple(p[:4]))
    return items


def list_campuses():
    """設定/ 底下有哪些院區。沒有這個資料夾就回空清單（走舊的根目錄設定）。"""
    if not os.path.isdir(CONFIG_DIR):
        return []
    return sorted(d for d in os.listdir(CONFIG_DIR)
                  if os.path.isdir(os.path.join(CONFIG_DIR, d)))


def config_path(campus, filename):
    """某院區的設定檔路徑。campus 為空就用根目錄那份。"""
    if campus:
        return os.path.join(CONFIG_DIR, campus, filename)
    return os.path.join(DATA_BASE_DIR, filename)


def load_settings():
    """讀本機設定（上次選的院區、上次用的模式）。"""
    out = {"院區": "", "模式": "one"}
    try:
        with open(LOCAL_SETTINGS, encoding="utf-8") as f:
            for line in f:
                k, _, v = line.strip().partition("=")
                if k in out and v:
                    out[k] = v
    except Exception:
        pass
    return out


def save_settings(**kw):
    """更新本機設定。只覆寫有傳進來的欄位。"""
    cur = load_settings()
    cur.update({k: v for k, v in kw.items() if v})
    try:
        with open(LOCAL_SETTINGS, "w", encoding="utf-8") as f:
            for k, v in cur.items():
                f.write(f"{k}={v}\n")
    except Exception:
        pass


def load_doctors(path=DOCTORS_FILE):
    """讀該院區的全部帳號。回傳 [(帳號, 顯示名稱), ...]

    代班不一定丟給均分池，也可能直接指定某位同事接手，
    所以下拉選單要看得到人。
    """
    items = []
    if not os.path.exists(path):
        return items
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            p = [x.strip() for x in next(csv.reader([line]))]
            if not p or p[0] == "帳號":
                continue
            items.append((p[0], p[1] if len(p) > 1 else p[0]))
    return items


def load_options(path=OPTIONS_FILE):
    """讀各欄位可以填什麼。回傳 {欄位: [值, ...]}，已按公版出現次數排好。

    來源是公版班表和工作時段設定（見 build_configs.py），
    比「目前畫面上出現過的值」完整 —— 畫面只查得到當天那幾列。
    """
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            p = [x.strip() for x in next(csv.reader([line]))]
            if len(p) < 2 or p[0] == "欄位":
                continue
            out.setdefault(p[0], []).append(p[1])
    return out


class Mapping:
    """一個院區的對照表：規則 + 「拿哪幾欄比對」 + 排除清單。

    比對欄位不能寫死。各院區決定代班對象的依據不一樣 ——
    新竹臺大看檢查項目（分類/類別/檢查室），
    癌醫看起來是看星期和時段（R01AM均分~R07AM均分）。
    所以由對照表自己宣告。
    """

    def __init__(self, rules=None, match_cols=None, excludes=None):
        self.rules = rules or []
        self.match_cols = tuple(match_cols or MATCH_COLS)
        self.excludes = excludes or []      # [(欄位, 值, 說明), ...]

    def __len__(self):
        return len(self.rules)


def load_mapping(path=MAPPING_FILE):
    """讀對照表。只讀主區段，忽略均分池區段。

    支援兩種 #! 指令（寫在註解裡，一般註解不受影響）：
        #!比對欄位=統計用分類碼,排班類別,檢查室名稱
        #!排除=檢查地點位址=竹東
    「排除」是院區專屬的業務規則（例如竹東的班在別的地方設定），
    以前寫死在程式裡，換院區就得改程式，所以搬到設定檔。
    """
    rules, match_cols, excludes = [], None, []
    if not os.path.exists(path):
        return Mapping()

    in_main = False
    with open(path, encoding="utf-8-sig") as f:
        for raw in f:
            line = raw.strip()
            if line.startswith("#!"):
                key, _, val = line[2:].partition("=")
                key, val = key.strip(), val.strip()
                if key == "比對欄位":
                    match_cols = [c.strip() for c in val.split(",") if c.strip()]
                elif key == "排除":
                    col, _, want = val.partition("=")
                    if col.strip() and want.strip():
                        excludes.append((col.strip(), want.strip()))
                continue
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in next(csv.reader([line]))]
            if parts[0] == "池帳號":
                in_main = False
                continue
            if match_cols and parts[0] == match_cols[0]:
                in_main = True          # 這是標題列
                continue
            if not match_cols and parts[0] == "統計用分類碼":
                in_main = True
                continue
            if not in_main or len(parts) < 4:
                continue
            while len(parts) < 6:
                parts.append("")
            cols = match_cols or list(MATCH_COLS)
            rule = {c: norm(parts[i]) for i, c in enumerate(cols)}
            rule.update({
                "代班帳號": norm(parts[len(cols)]),
                "顯示名稱": norm(parts[len(cols) + 1]) if len(parts) > len(cols) + 1 else "",
                "備註": norm(parts[len(cols) + 2]) if len(parts) > len(cols) + 2 else "",
            })
            rules.append(rule)
    return Mapping(rules, match_cols, excludes)


def match_rule(row, mapping):
    """由上往下找第一個對上的規則。回傳 (代班帳號 或 None, 規則 或 None)。

    對上規則但帳號留空 = 對照表刻意說「這種班交給人工」，
    這跟「完全沒有這條規則」是兩件事，所以規則本身也一起回傳。
    """
    for r in mapping.rules:
        for key in mapping.match_cols:
            want = r.get(key, "*")
            if want == "*":
                continue
            if norm(row.get(key)) != want:
                break
        else:
            return (r["代班帳號"] or None), r
    return None, None


def plan_for(row, mapping):
    """這一列該改成什麼。回傳 (帳號 或 None, 顯示名稱 或 None, 說明)。"""
    for col, want in mapping.excludes:
        if norm(row.get(col)) == want:
            return None, None, f"{want}（設定檔排除）"
    code, rule = match_rule(row, mapping)
    if code:
        return code, rule["顯示名稱"] or None, rule["備註"]
    if rule:
        return None, None, rule["備註"] or "對照表留空，交給人工"
    return None, None, "對照表沒有這種班別"


# ----------------------------------------------------------------------
# 介面
# ----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RIS 班表代班小工具")
        self.geometry("1180x760")

        self.ris = Ris()
        self.all_rows = []       # 讀進來的全部列
        self.shown_rows = []     # 篩選後顯示中的列
        self.settings = load_settings()
        self.campuses = list_campuses()
        self.campus = self.settings["院區"]
        if self.campuses and self.campus not in self.campuses:
            # 沒有記錄過就挑設定最完整的那個當預設 ——
            # 照字母排會選到還沒建規則的院區，開起來像壞掉的。
            self.campus = max(self.campuses,
                              key=lambda c: (len(load_mapping(config_path(c, MAPPING_FILE))),
                                             len(load_accounts(config_path(c, ACCOUNTS_FILE)))))
            save_settings(院區=self.campus)
        self.accounts = []
        self.rules = Mapping()
        self._load_config()
        self.q = queue.Queue()
        self.busy = False

        self._build()
        self.after(100, self._drain)

    # ---------- 設定 ----------
    def _load_config(self):
        """讀目前院區的帳號清單與對照表。"""
        self.accounts = load_accounts(config_path(self.campus, ACCOUNTS_FILE))
        self.doctors = load_doctors(config_path(self.campus, DOCTORS_FILE))
        self.options = load_options(config_path(self.campus, OPTIONS_FILE))
        self.rules = load_mapping(config_path(self.campus, MAPPING_FILE))

    def _value_choices(self):
        """「改成」下拉要列什麼，看現在選的是哪一欄。

        排班醫師：代班帳號排前面（那是最常用的），全院醫師接在後面。
        其他欄位：用 欄位選項.csv（來自公版班表和工作時段設定），
                 那是該院區實際用得到的完整清單。
                 目前畫面上出現、但清單裡沒有的值也補進去 ——
                 清單是靜態的，畫面才是現況，兩邊都要看得到。
        """
        col = self.col_var.get() if hasattr(self, "col_var") else DOCTOR_COL
        if col == DOCTOR_COL:
            out = [f"{a[0]}  ({a[1] or '?'})"
                   f"{'' if a[2].upper() == 'Y' else '  ※未驗證'}"
                   for a in self.accounts]
            seen = {a[0] for a in self.accounts}
            out += [f"{code}  ({display})" for code, display in self.doctors
                    if code not in seen]
            return out

        out = list(self.options.get(col, []))
        have = set(out)
        for extra in sorted({norm(r[col]) for r in self.all_rows if norm(r[col])}):
            if extra not in have:
                out.append(extra)
        return out

    def _refresh_choices(self):
        self.acct_cb["values"] = self._value_choices()
        self.acct_var.set("")

    def on_campus(self):
        """換院區：重讀設定，清掉已選的值（別院區的帳號在這裡無效）。"""
        self.campus = self.campus_var.get()
        save_settings(院區=self.campus)
        self._load_config()
        self._refresh_choices()
        self.log(f"切換到「{self.campus}」。")
        if not self.rules:
            self.mode.set("one")
            self.log("  這個院區沒有自動對照規則，用「全部改成」手動選。")
        self.apply_filter()

    def on_column(self):
        """換欄位：自動對照只認排班醫師，其他欄位一律手動。"""
        col = self.col_var.get()
        if col != DOCTOR_COL:
            self.mode.set("one")
            self.rb_auto.config(state="disabled")
            self.lbl_col.config(
                text=f"注意：這次會改「{col}」，不是排班醫師")
        else:
            self.rb_auto.config(state="normal")
            self.lbl_col.config(text="")
        self._refresh_choices()
        self.apply_filter()

    def on_mode(self):
        save_settings(模式=self.mode.get())
        self.apply_filter()

    # ---------- 版面 ----------
    def _build(self):
        pad = dict(padx=6, pady=4)

        # 第零區：院區
        if self.campuses:
            cf = ttk.LabelFrame(self, text="0. 這批班是哪個院區的")
            cf.pack(fill="x", **pad)
            self.campus_var = tk.StringVar(value=self.campus)
            cb = ttk.Combobox(cf, textvariable=self.campus_var,
                              values=self.campuses, width=32, state="readonly")
            cb.pack(side="left", padx=8, pady=6)
            cb.bind("<<ComboboxSelected>>", lambda e: self.on_campus())
            ttk.Label(cf, text="選錯院區＝用錯對照表，改之前確認一下",
                      foreground="#b00").pack(side="right", padx=10)

        # 第一區：讀取
        top = ttk.LabelFrame(self, text="1. 先在 RIS 查詢出資料，再按這裡")
        top.pack(fill="x", **pad)
        self.btn_load = ttk.Button(top, text="讀取畫面表格", command=self.on_load)
        self.btn_load.pack(side="left", padx=6, pady=6)
        self.lbl_status = ttk.Label(top, text="尚未讀取")
        self.lbl_status.pack(side="left", padx=10)

        # 第二區：篩選
        mid = ttk.LabelFrame(self, text="2. 篩選要改的列")
        mid.pack(fill="x", **pad)
        self.fvars = {}
        for i, col in enumerate(FILTERS):
            r, c = divmod(i, 3)
            box = ttk.Frame(mid)
            box.grid(row=r, column=c, sticky="w", padx=8, pady=3)
            ttk.Label(box, text=col, width=10).pack(side="left")
            v = tk.StringVar(value="（全部）")
            cb = ttk.Combobox(box, textvariable=v, width=24, state="readonly")
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())
            self.fvars[col] = (v, cb)
        ttk.Button(mid, text="清除篩選", command=self.clear_filter)\
            .grid(row=1, column=3, padx=10)

        # 第三區：清單
        listf = ttk.LabelFrame(self, text="3. 確認清單（可用 Ctrl / Shift 只選取部分列）")
        listf.pack(fill="both", expand=True, **pad)
        self.tree = ttk.Treeview(listf, columns=["列"] + SHOW + ["將改成"],
                                 show="headings", selectmode="extended")
        self.tree.heading("列", text="列")
        self.tree.column("列", width=45, anchor="center")
        widths = {"起始日期": 95, "排班時段": 150, "統計用分類碼": 90,
                  "排班類別": 110, "檢查室名稱": 150, "門急住限制": 80,
                  "檢查地點位址": 90, "排班醫師": 140, "將改成": 190}
        for c in SHOW + ["將改成"]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths.get(c, 100), anchor="w")
        # 對照表接不到的列用灰字，一眼就看得出哪些要自己改
        self.tree.tag_configure("manual", foreground="#999")
        sb = ttk.Scrollbar(listf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 第四區：執行
        bot = ttk.LabelFrame(self, text="4. 改哪一欄、改成什麼")
        bot.pack(fill="x", **pad)

        r0 = ttk.Frame(bot); r0.pack(fill="x")
        ttk.Label(r0, text="改哪一欄").pack(side="left", padx=(8, 2))
        self.col_var = tk.StringVar(value=DOCTOR_COL)
        col_cb = ttk.Combobox(r0, textvariable=self.col_var, values=EDITABLE,
                              width=14, state="readonly")
        col_cb.pack(side="left", padx=4)
        col_cb.bind("<<ComboboxSelected>>", lambda e: self.on_column())
        self.lbl_col = ttk.Label(r0, text="", foreground="#b00")
        self.lbl_col.pack(side="left", padx=10)

        # 預設是手動 —— 自動對照要有人維護過對照表才可信，
        # 沒建規則的院區用自動只會整片顯示「手動」，反而像壞掉。
        # 用過一次會記在本機設定，下次直接沿用。
        r2 = ttk.Frame(bot); r2.pack(fill="x")
        self.mode = tk.StringVar(value=self.settings.get("模式", "one"))
        if not self.rules:
            self.mode.set("one")
        ttk.Radiobutton(r2, text="全部改成", variable=self.mode, value="one",
                        command=self.on_mode).pack(side="left", padx=(8, 2))
        self.acct_var = tk.StringVar()
        self.acct_cb = ttk.Combobox(r2, textvariable=self.acct_var,
                                    values=self._value_choices(), width=42)
        self.acct_cb.pack(side="left", padx=4)
        self.acct_cb.bind("<<ComboboxSelected>>", lambda e: self._pick_one())
        ttk.Label(r2, text="（也可以直接打）").pack(side="left")

        r1 = ttk.Frame(bot); r1.pack(fill="x")
        self.rb_auto = ttk.Radiobutton(
            r1, text="自動對照（照對照表，每列各自決定；只適用排班醫師）",
            variable=self.mode, value="auto", command=self.on_mode)
        self.rb_auto.pack(side="left", padx=(8, 4))

        r3 = ttk.Frame(bot); r3.pack(fill="x")
        self.only_one = tk.BooleanVar(value=True)
        ttk.Checkbutton(r3, text="先只改一列試試", variable=self.only_one)\
            .pack(side="left", padx=16)

        self.btn_go = ttk.Button(r3, text="開始修改", command=self.on_apply)
        self.btn_go.pack(side="left", padx=10, pady=6)

        ttk.Label(r3, text="程式不會存檔，改完請自己回 RIS 檢查並按存檔",
                  foreground="#b00").pack(side="left", padx=10)

        # 訊息
        logf = ttk.LabelFrame(self, text="訊息")
        logf.pack(fill="both", **pad)
        self.txt = tk.Text(logf, height=10, wrap="word")
        self.txt.pack(fill="both", expand=True, side="left")
        sb2 = ttk.Scrollbar(logf, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscroll=sb2.set)
        sb2.pack(side="right", fill="y")

        self.log("提醒：不要點 RIS 表格最左邊的列首。那會選取整列，"
                 "存檔時會變成刪除該列。")

    # ---------- 小工具 ----------
    def log(self, s):
        self.txt.insert("end", f"{datetime.now():%H:%M:%S}  {s}\n")
        self.txt.see("end")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {s}\n")

    def _drain(self):
        """背景執行緒透過 queue 回報訊息，這裡負責更新畫面。"""
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "status":
                    self.lbl_status.config(text=payload)
                elif kind == "rows":
                    self.all_rows = payload
                    self._fill_filters()
                    self.apply_filter()
                    self.lbl_status.config(text=f"讀到 {len(payload)} 列")
                elif kind == "popup":
                    title, body, ok = payload
                    (messagebox.showinfo if ok else messagebox.showwarning)(title, body)
                elif kind == "done":
                    self.busy = False
                    self.btn_load.config(state="normal")
                    self.btn_go.config(state="normal")
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _run(self, fn):
        if self.busy:
            return
        self.busy = True
        self.btn_load.config(state="disabled")
        self.btn_go.config(state="disabled")

        def wrap():
            try:
                fn()
            except Exception as e:
                self.q.put(("log", f"!! 發生錯誤: {type(e).__name__}: {e}"))
                self.q.put(("log", traceback.format_exc()))
            finally:
                self.q.put(("done", None))

        threading.Thread(target=wrap, daemon=True).start()

    # ---------- 讀取 ----------
    def on_load(self):
        def job():
            self.q.put(("status", "連線中..."))
            self.ris.connect()
            self.q.put(("status", "讀取表格中..."))
            t0 = time.time()
            rows = self.ris.read_rows(
                progress=lambda i, n: self.q.put(("status", f"讀取中 {i}/{n}")))
            self.q.put(("rows", rows))
            self.q.put(("log", f"讀到 {len(rows)} 列。（{time.time() - t0:.1f} 秒）"))
            if not rows:
                self.q.put(("log", "表格是空的 —— 請確認 RIS 已經查詢出資料。"))
        self._run(job)

    def _fill_filters(self):
        for col, (var, cb) in self.fvars.items():
            vals = sorted({norm(r[col]) for r in self.all_rows if norm(r[col])})
            cb["values"] = ["（全部）"] + vals
            var.set("（全部）")

    def clear_filter(self):
        for _col, (var, _cb) in self.fvars.items():
            var.set("（全部）")
        self.apply_filter()

    def _pick_one(self):
        """在下拉選了值 = 想用「全部改成同一個」，順手把模式切過去。"""
        self.mode.set("one")
        self.apply_filter()

    def target_col(self):
        return self.col_var.get() if hasattr(self, "col_var") else DOCTOR_COL

    def plan_of(self, row):
        """這一列會被改成什麼。回傳 (要打的值 或 None, 預期顯示 或 None, 說明)。"""
        # 設定檔宣告要排除的（例如竹東的班在別的地方設定），
        # 不管哪種模式都不碰 —— 手動模式也一樣，免得整批套下去掃到。
        for col, want in self.rules.excludes:
            if norm(row.get(col)) == want:
                return None, None, f"{want}（設定檔排除）"
        if self.mode.get() == "auto":
            return plan_for(row, self.rules)
        code, display = self._target_code()
        if not code:
            return None, None, "還沒選要改成什麼"
        return code, display, ""

    def apply_filter(self):
        rows = []
        for r in self.all_rows:
            ok = True
            for col, (var, _cb) in self.fvars.items():
                want = var.get()
                if want != "（全部）" and norm(r[col]) != want:
                    ok = False
                    break
            if ok:
                rows.append(r)
        self.shown_rows = rows

        self.tree.delete(*self.tree.get_children())
        auto_ok = 0
        for r in rows:
            code, display, why = self.plan_of(r)
            if code:
                auto_ok += 1
                target = display or code
                tags = ()
            else:
                target = f"手動：{why}" if why else "手動"
                tags = ("manual",)
            self.tree.insert("", "end", iid=str(r["_row"]), tags=tags,
                             values=[r["_row"]] + [norm(r[c]) for c in SHOW]
                                    + [target])

        msg = f"讀到 {len(self.all_rows)} 列，符合條件 {len(rows)} 列"
        if rows:
            msg += f"，其中 {auto_ok} 列可自動改、{len(rows) - auto_ok} 列要手動"
        self.lbl_status.config(text=msg)

    # ---------- 修改 ----------
    def _target_code(self):
        """從「改成」下拉解析出 (要打的值, 預期顯示)。

        排班醫師的選項長這樣：「CCTAOFF  (CCTAOFF均分)」，
        要打的是帳號、預期顯示是括號裡那串。
        其他欄位沒有這種帳號/顯示的分別，選什麼就打什麼。
        """
        raw = self.acct_var.get().strip()
        if not raw:
            return None, None
        if self.target_col() != DOCTOR_COL:
            return raw, raw
        code = raw.split()[0]
        display = None
        for a in self.accounts:
            if a[0] == code:
                display = a[1] or None
        if display is None:
            for dcode, ddisp in self.doctors:
                if dcode == code:
                    display = ddisp
        return code, display

    def on_apply(self):
        col = self.target_col()
        if self.mode.get() == "one" and not self._target_code()[0]:
            messagebox.showwarning("還沒選", f"請先選（或輸入）「{col}」要改成什麼。")
            return

        sel = self.tree.selection()
        if sel:
            picked = [r for r in self.shown_rows if str(r["_row"]) in sel]
            scope = "選取的"
        else:
            picked = list(self.shown_rows)
            scope = "篩選出來的全部"
        if not picked:
            messagebox.showwarning("沒有列", "清單裡沒有可以修改的列。")
            return

        # 每一列各自算目標帳號；算不出來的留給人工，不猜
        targets = []          # [(列, 帳號, 顯示名稱)]
        manual = []           # [(列, 原因)]
        for r in picked:
            code, display, why = self.plan_of(r)
            if code:
                targets.append((r, code, display))
            else:
                manual.append((r, why))

        if not targets:
            messagebox.showwarning(
                "沒有可自動處理的列",
                f"選到的 {len(picked)} 列，對照表都接不到，要自己改。\n\n"
                + "\n".join(f"  列{r['_row']}  {norm(r['統計用分類碼'])} "
                            f"{norm(r['排班類別'])} —— {why}"
                            for r, why in manual[:15]))
            return

        if self.only_one.get():
            targets = targets[:1]

        by_code = {}
        for _r, code, _d in targets:
            by_code[code] = by_code.get(code, 0) + 1
        summary = "\n".join(f"    {c}  ×{n} 列" for c, n in sorted(by_code.items()))

        preview = "\n".join(
            f"  列{r['_row']:>3}  {norm(r['統計用分類碼']):5} {norm(r['排班類別']):12} "
            f"{norm(r['排班醫師'])}  ->  {d or c}"
            for r, c, d in targets[:15])
        more = f"\n  ...另外還有 {len(targets) - 15} 列" if len(targets) > 15 else ""

        skipped = ""
        if manual:
            skipped = (f"\n\n※ 另外 {len(manual)} 列對照表接不到，"
                       "這次不會動，要你自己改：\n"
                       + "\n".join(f"    列{r['_row']}  {norm(r['統計用分類碼'])} "
                                   f"{norm(r['排班類別'])} —— {why}"
                                   for r, why in manual[:10]))

        msg = (f"要改{scope}其中 {len(targets)} 列：\n\n"
               f"{summary}\n\n{preview}{more}{skipped}\n\n"
               "程式不會存檔，改完你要自己回 RIS 檢查並按存檔。\n"
               "確定要執行嗎？")
        if not messagebox.askyesno("再確認一次", msg):
            self.log("已取消，什麼都沒改。")
            return

        def job():
            self.q.put(("log", "=" * 60))
            self.q.put(("log", f"開始修改 {len(targets)} 列的「{col}」  "
                               + "、".join(f"{c}×{n}" for c, n
                                           in sorted(by_code.items()))))
            t_start = time.time()
            done = 0
            stopped = None
            written = []      # [(列號, 期待值)]，最後拿來全部核對
            attempted = set()  # 動過手的列。失敗的列也算，免得收尾核對
                               # 把「試過但失敗」誤報成「根本不該被動到」
            for r, code, display in targets:
                idx = r["_row"]
                t_row = time.time()
                # 改之前重新讀一次，確認這一列還是原本那一列
                now = self.ris.read_row(idx)
                if now is None:
                    self.q.put(("log", f"  X 第 {idx} 列讀不到了，中止。"))
                    stopped = f"第 {idx} 列讀不到了"
                    break
                bad = [c for c in ("起始日期", "排班時段", "統計用分類碼",
                                   "排班類別", "排班醫師")
                       if norm(now.get(c)) != norm(r.get(c))]
                if bad:
                    self.q.put(("log", f"  X 第 {idx} 列內容變了（{', '.join(bad)}），"
                                       f"中止以免改錯。請重新讀取表格。"))
                    stopped = (f"第 {idx} 列的內容跟讀取時不一樣了"
                               f"（{', '.join(bad)}）。\n"
                               "為避免改錯已中止，請重新讀取表格。")
                    break

                # 每一列都要走完「點進去 -> 打字 -> Enter」，一步都不能省。
                t_goto = time.time()
                moved, how = self.ris.goto_row(idx, col)
                t_goto = time.time() - t_goto
                if not moved:
                    self.q.put(("log", f"  X 第 {idx} 列：{how}，中止（不在不確定的位置寫入）。"))
                    stopped = (f"游標移不到第 {idx} 列：{how}\n\n"
                               "為避免寫錯格子已中止。")
                    break

                attempted.add(idx)
                ok, after, why = self.ris.write_cell(idx, col, code, display)
                if ok:
                    done += 1
                    written.append((idx, display or code))
                    ty, au, cm, n_ent = self.ris.last_timing
                    self.q.put(("log", f"  V 第 {idx} 列  "
                                       f"{norm(r[col])} -> {after}"
                                       f"  ({time.time() - t_row:.1f}s ="
                                       f" 點{t_goto:.1f} 打{ty:.1f}"
                                       f" 補{au:.1f} 提交{cm:.1f}/{n_ent}次Enter)"))
                    # 2026-08-10 效能診斷：只加這兩行，不動上面那行的格式。
                    # 診斷完不需要了整段刪掉即可，其他都不受影響。
                    (t_fnd, t_eon, t_scr, t_foc, t_clk,
                     t_ew, n_clk) = self.ris.last_goto_detail
                    t_e1, t_e2, t_w1 = self.ris.last_commit_detail
                    n_ep, t_ep, t_epmax = self.ris.last_editor_poll_detail
                    n_wp, t_wp, t_wpmax = self.ris.last_wait_value_detail
                    self.q.put(("log", f"      debug 點=[找格子{t_fnd:.2f} "
                                       f"早期editor檢查{t_eon:.2f} 捲{t_scr:.2f} "
                                       f"焦{t_foc:.2f} 擊{t_clk:.2f} "
                                       f"等編輯器{t_ew:.2f}/{n_clk}次點擊] "
                                       f"提交=[Enter1={t_e1:.2f} Enter2={t_e2:.2f} "
                                       f"等格子{t_w1:.2f}]"))
                    self.q.put(("log", f"      debug poll細節: "
                                       f"等編輯器共問{n_ep}次/總{t_ep:.2f}s/單次最久{t_epmax:.2f}s"
                                       f"　等格子共問{n_wp}次/總{t_wp:.2f}s/單次最久{t_wpmax:.2f}s"))
                else:
                    self.q.put(("log", f"  X 第 {idx} 列（打 {code}）：{why}，中止。"))
                    stopped = (f"第 {idx} 列打進 {code} 的時候：\n\n{why}\n\n"
                               f"帳號 {code} 可能不正確。\n"
                               "請回 RIS 檢查，需要的話重新查詢並選「否」還原。")
                    break

            # 收尾核對：重讀整張表，逐列比對。
            # 不能只檢查改過的那幾列 —— 提交用的方向鍵是打在下拉式方塊上的，
            # 萬一改到隔壁那些「不該動」的列，只看目標列永遠看不出來。
            if done:
                self.q.put(("log", "  重讀整張表核對中..."))
                want = {r["_row"]: norm(r[col]) for r in self.all_rows}
                want.update({i: norm(v) for i, v in written})

                try:
                    fresh = self.ris.read_rows()
                except Exception as e:
                    fresh = None
                    self.q.put(("log", f"  !! 重讀失敗（{type(e).__name__}），"
                                       "請自己在畫面上逐列確認再決定要不要存檔。"))

                if fresh is not None:
                    wrong, strays = [], []
                    for x in fresh:
                        i, got = x["_row"], norm(x[col])
                        if i not in want:
                            continue
                        if got == want[i] or got.startswith(want[i]):
                            continue
                        (wrong if i in attempted else strays).append(
                            (i, got, want[i]))

                    for i, got, exp in wrong:
                        self.q.put(("log", f"     第 {i} 列：現在是 {got!r}，預期 {exp!r}"))
                    for i, got, exp in strays:
                        self.q.put(("log", f"  !! 第 {i} 列根本不該被動到，"
                                           f"卻從 {exp!r} 變成 {got!r}"))

                    if strays:
                        stopped = (f"有 {len(strays)} 列是這次「不該動」的，卻被改掉了。\n\n"
                                   "請不要存檔。回 RIS 重新查詢，提示選「否」還原，"
                                   "然後把訊息告訴我。")
                    elif wrong:
                        stopped = (f"改完後核對發現 {len(wrong)} 列的值不對。\n\n"
                                   "建議不要存檔，回 RIS 重新查詢並選「否」還原，"
                                   "然後把訊息記下來。")
                    else:
                        self.q.put(("log", f"  核對通過：{done} 列已改，"
                                           "其他列都沒被動到。"))

            self.q.put(("log", f"完成 {done} / {len(targets)} 列"
                               f"（共 {time.time() - t_start:.1f} 秒）"))
            if done:
                self.q.put(("log", ">>> 請回 RIS 檢查，沒問題再自己按存檔。"))
                self.q.put(("log", ">>> 要放棄的話：重新查詢，提示選「否」。"))
            self.q.put(("log", "=" * 60))

            # 跳視窗提醒，免得改完沒注意到、忘了存檔
            if stopped:
                self.q.put(("popup", (
                    "中途停止",
                    f"已改了 {done} 列，然後停在這裡：\n\n{stopped}\n\n"
                    "程式沒有存檔。",
                    False)))
            elif done:
                self.q.put(("popup", (
                    "改好了，請去存檔",
                    f"已經改好 {done} 列：\n"
                    + "\n".join(f"    {c}  ×{n} 列"
                                for c, n in sorted(by_code.items())) + "\n\n"
                    "接下來請你自己做：\n"
                    "  1. 切回 RIS 看一下改的內容對不對\n"
                    "  2. 沒問題就按「存檔」\n"
                    "  3. 想放棄的話，重新查詢一次，跳出提示時選「否」\n\n"
                    "※ 程式沒有存檔，現在資料庫裡還是原本的內容。\n"
                    "※ 在你決定之前，不要關視窗或重新查詢。",
                    True)))

        self._run(job)


if __name__ == "__main__":
    App().mainloop()
