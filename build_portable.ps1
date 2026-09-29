# 做「免安裝版」：python.org 官方的 embeddable Python + ris_gui.py + 設定/
#
# 為什麼不用 exe：分院同事的電腦上，Nuitka 編出來的 exe 都被防毒擋掉
# （沒簽章的新 exe）。python.org 官方的 python.exe / pythonw.exe 有 Python
# Software Foundation 的數位簽章，程式本身則是純文字的 .py 檔。
#
# 產出：portable_build\RIS代班工具\  整個資料夾壓成 zip 發出去即可。
#   RIS代班工具.bat     雙擊啟動
#   建立桌面捷徑.bat    在那台電腦的桌面建一個捷徑（捷徑要絕對路徑，只能到現場建；
#                       實際工作在 make_shortcut.py）
#   ris_gui.py、設定\   程式和設定檔，改版時只要換 ris_gui.py
#   python\             免安裝 Python（含 tkinter 和 pywinauto 等套件）
#
# 執行： powershell -ExecutionPolicy Bypass -File build_portable.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

# 版本要跟本機 Python 一樣，因為 tkinter 是從本機那一份複製過去的
$ver = (python -c "import platform; print(platform.python_version())").Trim()
$localPy = (python -c "import sys; print(sys.base_prefix)").Trim()
$tag = "python" + ($ver.Split(".")[0..1] -join "")          # 例：python312

$out = Join-Path $here "portable_build\RIS代班工具"
$py = Join-Path $out "python"
$cache = Join-Path $here "portable_build\_cache"
New-Item -ItemType Directory -Force $cache | Out-Null
if (Test-Path $out) { Remove-Item -Recurse -Force $out }
New-Item -ItemType Directory -Force $py | Out-Null

# 1. 官方 embeddable Python
$zipName = "python-$ver-embed-amd64.zip"
$zip = Join-Path $cache $zipName
if (-not (Test-Path $zip)) {
    Write-Host "下載 $zipName ..."
    Invoke-WebRequest "https://www.python.org/ftp/python/$ver/$zipName" -OutFile $zip
}
Expand-Archive $zip -DestinationPath $py

# 2. tkinter（embeddable 版沒有附，GUI 需要）
Copy-Item -Recurse (Join-Path $localPy "Lib\tkinter") (Join-Path $py "tkinter")
# zlib1.dll 是 tcl86t.dll 要的，漏了會「DLL load failed while importing _tkinter」
foreach ($f in "_tkinter.pyd", "tcl86t.dll", "tk86t.dll", "zlib1.dll") {
    Copy-Item (Join-Path $localPy "DLLs\$f") $py
}
New-Item -ItemType Directory -Force (Join-Path $py "tcl") | Out-Null
foreach ($d in "tcl8.6", "tk8.6") {
    Copy-Item -Recurse (Join-Path $localPy "tcl\$d") (Join-Path $py "tcl\$d")
}
# tcl8.6 的 init.tcl 會去找 tcl8 底下的模組
Copy-Item -Recurse (Join-Path $localPy "tcl\tcl8") (Join-Path $py "tcl\tcl8")

# 3. 套件：版本鎖死成本機實測過的那一組
$site = Join-Path $py "Lib\site-packages"
python -m pip install --quiet --disable-pip-version-check --no-deps `
    --target $site --only-binary=:all: `
    --platform win_amd64 --python-version $ver --implementation cp `
    pywinauto==0.6.9 comtypes==1.4.16 pywin32==312 six==1.17.0
if ($LASTEXITCODE -ne 0) { throw "pip 安裝套件失敗" }

# 4. 讓 embeddable Python 看得到 site-packages（它預設只看 ._pth 列的路徑，
#    而且不跑 site，pywin32 靠 .pth 設定 DLL 路徑，所以 import site 要打開）
$pth = Join-Path $py "$tag._pth"
@("$tag.zip", ".", "Lib\site-packages", "import site") |
    Set-Content -Encoding ascii $pth

# 4a. 刪掉 pywin32 附的、這個工具絕對用不到的執行檔（沒有數位簽章，
#     防毒最容易盯上 exe）。.pyd 模組全部保留 —— 有些要到真的改班時才載入，
#     刪錯了會在同事那邊才爆。2026-09-29 實測讀表+GUI 用到的是：
#     win32api/win32gui/win32process/win32event/win32ui/_win32sysloader/
#     pythoncom/pywintypes/shell(建捷徑)。
foreach ($rel in "bin", "isapi", "pythonwin\Pythonwin.exe", "win32\pythonservice.exe") {
    $p = Join-Path $site $rel
    if (Test-Path $p) { Remove-Item -Recurse -Force $p }
}

# 4b. 先用包裡的 Python 把套件編譯成 .pyc：同事第一次開比較快，
#     也不會跳 pywinauto 原始碼裡的 SyntaxWarning
#     用 unchecked-hash：zip 裡的時間只精確到 2 秒，解壓縮後時間對不上，
#     預設（看時間）的 .pyc 會被當成過期而重新編譯
& (Join-Path $py "python.exe") -W ignore -m compileall -q -f --invalidation-mode unchecked-hash $site | Out-Null

# 5. 程式和設定檔
Copy-Item (Join-Path $here "ris_gui.py") $out
Copy-Item -Recurse (Join-Path $here "設定") (Join-Path $out "設定")

# 6. 啟動檔。bat 裡只放英文和 %~dp0（這個 bat 所在的資料夾）——
#    bat 裡的中文會隨命令列編碼變亂碼（2026-09-29 實測，捷徑名稱變成
#    「RIS?N?Z?u??」建立失敗），要中文的事交給 Python 做。
#    整包搬到哪裡都能跑。
$launch = @"
@echo off
start "" "%~dp0python\pythonw.exe" "%~dp0ris_gui.py"
"@
$mkLink = @"
@echo off
"%~dp0python\python.exe" "%~dp0make_shortcut.py"
pause
"@
foreach ($pair in @(@("RIS代班工具.bat", $launch), @("建立桌面捷徑.bat", $mkLink))) {
    $text = $pair[1].Replace("`r`n", "`n").Replace("`n", "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $out $pair[0]), $text,
                                   [System.Text.Encoding]::ASCII)
}
Copy-Item (Join-Path $here "portable_make_shortcut.py") (Join-Path $out "make_shortcut.py")

# 7. 壓成 zip 方便傳
$outZip = Join-Path $here "portable_build\RIS代班工具_免安裝版.zip"
if (Test-Path $outZip) { Remove-Item -Force $outZip }
# 不用 Compress-Archive：PowerShell 5.1 壓出來的中文檔名，別台電腦解壓縮可能變亂碼。
# Python 的 zipfile 遇到非 ASCII 檔名會標 UTF-8，Windows 檔案總管認得。
python -c "import shutil, sys; shutil.make_archive(sys.argv[1][:-4], 'zip', sys.argv[2], 'RIS代班工具')" $outZip (Split-Path $out)
if ($LASTEXITCODE -ne 0) { throw "壓縮失敗" }

$size = "{0:N1}" -f ((Get-Item $outZip).Length / 1MB)
Write-Host ""
Write-Host "完成：$out"
Write-Host "      $outZip （$size MB）"
Write-Host ""
Write-Host "發給分院前："
Write-Host "  1. 設定\ 底下只留對方那個院區的資料夾，其他刪掉（要的話）"
Write-Host "  2. 對方解壓縮後雙擊「建立桌面捷徑.bat」一次，之後用桌面捷徑開"
Write-Host "  3. 第一次真的改班照規矩勾「先只改一列試試」，存檔後重新查詢確認"
