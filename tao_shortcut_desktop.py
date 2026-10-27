# -*- coding: utf-8 -*-
"""
Script tự động tạo Shortcut ngoài màn hình Desktop với Logo chuẩn Viện Kiểm Sát.
"""
import os
import sys

def tao_shortcut():
    try:
        import win32com.client
    except ImportError:
        pass

    desktop = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
    base_dir = os.path.dirname(os.path.abspath(__file__))
    app_script = os.path.join(base_dir, 'app_vks.py')
    icon_path = os.path.join(base_dir, 'assets', 'logo_vks.ico')
    
    # Tìm pythonw.exe
    python_dir = os.path.dirname(sys.executable)
    pythonw = os.path.join(python_dir, 'pythonw.exe')
    if not os.path.exists(pythonw):
        pythonw = sys.executable

    shortcut_name = "Phần mềm Tổng hợp Án VKS.lnk"
    shortcut_path = os.path.join(desktop, shortcut_name)

    ps_script = f'''
    $wsh = New-Object -ComObject WScript.Shell
    $s = $wsh.CreateShortcut('{shortcut_path}')
    $s.TargetPath = '{pythonw}'
    $s.Arguments = '"{app_script}"'
    $s.WorkingDirectory = '{base_dir}'
    $s.IconLocation = '{icon_path},0'
    $s.Description = 'Hệ thống Tổng hợp & Giám sát việc lập hồ sơ kiểm sát điện tử - VKSND TP.HCM'
    $s.Save()
    '''

    import subprocess
    cmd = ["powershell", "-NoProfile", "-Command", ps_script]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode == 0:
        print(f"Đã tạo shortcut thành công tại: {shortcut_path}")
        return True
    else:
        print("Lỗi tạo shortcut:", res.stderr)
        return False

if __name__ == "__main__":
    tao_shortcut()
