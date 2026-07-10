# UTF-8
$Cwd = (Get-Location).Path
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$ShortcutPath = Join-Path $DesktopPath "Барсик - Умная сверка.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = Join-Path $Cwd "run.bat"
$Shortcut.WorkingDirectory = $Cwd
$Shortcut.IconLocation = Join-Path $Cwd "assets\barsik_logo.ico"
$Shortcut.Save()

Write-Host "Ярлык успешно создан на Рабочем столе!"
