"""在這台電腦的桌面建立「RIS代班工具」捷徑。免安裝版專用（build_portable.ps1 會複製成 make_shortcut.py）。

捷徑要寫絕對路徑，所以只能在同事的電腦上、解壓縮完之後才建。
不寫在 bat 裡的原因：bat 裡的中文會隨命令列編碼變成亂碼（2026-09-29 實測，
捷徑名稱變成「RIS?N?Z?u??」建立失敗）。
"""

import os
import sys

import win32com.client
from win32com.shell import shell, shellcon

here = os.path.dirname(os.path.abspath(__file__))
pythonw = os.path.join(here, "python", "pythonw.exe")
script = os.path.join(here, "ris_gui.py")

try:
    desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, None, 0)
    link = os.path.join(desktop, "RIS代班工具.lnk")
    s = win32com.client.Dispatch("WScript.Shell").CreateShortcut(link)
    s.TargetPath = pythonw
    s.Arguments = f'"{script}"'
    s.WorkingDirectory = here
    s.IconLocation = pythonw + ",0"
    s.Description = "RIS 醫師休假代班工具"
    s.Save()
except Exception as e:
    print(f"建立失敗：{e}")
    sys.exit(1)

print(f"已在桌面建立捷徑：{link}")
print("之後用桌面上的「RIS代班工具」開啟即可。")
