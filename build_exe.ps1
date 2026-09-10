# 把 ris_gui.py 用 Nuitka 編成免裝 Python 的 exe。
#
# 用法：改完 ris_gui.py 之後，執行
#   .\build_exe.ps1                 # 兩種都包（預設）
#   .\build_exe.ps1 -Mode standalone   # 只包資料夾版
#   .\build_exe.ps1 -Mode onefile      # 只包單一 exe 版
#
# 2026-08-14 踩過的坑：如果跳出「因為這個系統上停用指令碼」之類的錯誤，
# 是 PowerShell 執行政策擋住直接執行 .ps1，換成：
#   powershell -ExecutionPolicy Bypass -File build_exe.ps1
#
# 兩種版本怎麼選（2026-08-14 討論過的結論）：
#   standalone（資料夾）：執行時不解壓縮，防毒誤判風險較低，推薦優先用這個。
#     產出：nuitka_build\ris_gui.dist\ris_gui.exe（連同整個資料夾一起帶走，
#     或直接分發 nuitka_build\ris_gui_standalone.zip）。
#   onefile（單一 exe）：只有一個檔案方便分發，但執行時會先解壓縮到
#     %TEMP%，跟 PyInstaller 的 onefile 是同一種行為模式，比較容易被
#     防毒軟體誤判——只是「不確定是不是真的會被擋」時的備案，不是首選。
#     產出：nuitka_build_onefile\ris_gui.exe（單一檔案，直接帶走就好）。
#
# 2026-08-14 踩過的坑：comtypes 動態產生 COM 包裝程式碼是用執行期才
# import 的方式（importlib.import_module），Nuitka 靜態分析看不到，
# 預設只會打包看得到的部分，導致一啟動就是
# ModuleNotFoundError: No module named 'comtypes.stream'。
# 解法是 --include-package=comtypes 強制把整個套件（含它用不到的
# comtypes.test 之類）都打包進去，副作用只是檔案變大、編譯變久。

param(
    [ValidateSet("both", "standalone", "onefile")]
    [string]$Mode = "both"
)

$ErrorActionPreference = "Stop"

$commonArgs = @(
    "--windows-console-mode=disable"
    "--enable-plugin=tk-inter"
    "--include-package=comtypes"
    "--include-data-dir=設定=設定"
    "--assume-yes-for-downloads"
)

function Build-Standalone {
    Write-Host "=== 編 standalone（資料夾）版 ===" -ForegroundColor Cyan
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue nuitka_build

    python -m nuitka --standalone --output-dir=nuitka_build @commonArgs ris_gui.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host "standalone 編譯失敗，看上面的錯誤訊息。" -ForegroundColor Red
        exit 1
    }

    Write-Host "壓成 zip 方便分發..."
    Remove-Item -Force -ErrorAction SilentlyContinue nuitka_build\ris_gui_standalone.zip
    Compress-Archive -Path nuitka_build\ris_gui.dist -DestinationPath nuitka_build\ris_gui_standalone.zip -Force

    Write-Host "standalone 完成：nuitka_build\ris_gui.dist\ris_gui.exe" -ForegroundColor Green
    Write-Host "              壓縮檔：nuitka_build\ris_gui_standalone.zip" -ForegroundColor Green
}

function Build-Onefile {
    Write-Host "=== 編 onefile（單一 exe）版 ===" -ForegroundColor Cyan
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue nuitka_build_onefile

    python -m nuitka --onefile --output-dir=nuitka_build_onefile @commonArgs ris_gui.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host "onefile 編譯失敗，看上面的錯誤訊息。" -ForegroundColor Red
        exit 1
    }

    Write-Host "onefile 完成：nuitka_build_onefile\ris_gui.exe" -ForegroundColor Green
}

switch ($Mode) {
    "standalone" { Build-Standalone }
    "onefile"    { Build-Onefile }
    "both"       { Build-Standalone; Build-Onefile }
}

Write-Host ""
Write-Host "分發給別的分院前，記得："
Write-Host "  1. 資料夾版的話，nuitka_build\ris_gui.dist\設定\ 底下只留對方那個院區的資料夾，其他刪掉"
Write-Host "     （onefile 版的設定檔是包進單一 exe 裡的，沒辦法事後刪，要分院區發就要分開編）"
Write-Host "  2. 對方拿到之後，第一次用照平常規矩：開著 RIS、勾「先只改一列試試」，實測一次寫入再信任"
