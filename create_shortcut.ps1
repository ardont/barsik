$s = New-Object -ComObject WScript.Shell
$sc = $s.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Барсик - Умная сверка.lnk'))
$sc.TargetPath = "$PSScriptRoot\run.bat"
$sc.WorkingDirectory = "$PSScriptRoot"
$sc.IconLocation = [System.IO.Path]::Combine($PSScriptRoot, 'assets', 'barsik_logo.ico')
$sc.Save()
