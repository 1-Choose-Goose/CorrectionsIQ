from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    main_py = project_dir / "main.py"
    icon_path = project_dir / "assets" / "app_icon.ico"
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    target = pythonw if pythonw.exists() else Path(sys.executable)

    script = f"""
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'CorrectionsIQ.lnk'
if (Test-Path $shortcutPath) {{
    Remove-Item -LiteralPath $shortcutPath -Force
}}
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = {ps_quote(str(target))}
$shortcut.Arguments = {ps_quote(str(main_py))}
$shortcut.WorkingDirectory = {ps_quote(str(project_dir))}
$shortcut.IconLocation = {ps_quote(str(icon_path) + ',0')}
$shortcut.Description = 'CorrectionsIQ'
$shortcut.Save()
Write-Output $shortcutPath
"""
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        text=True,
        capture_output=True,
    )
    print(completed.stdout.strip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
