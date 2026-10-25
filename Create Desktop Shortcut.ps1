# Creates a "Books to PDF" shortcut on the desktop that runs the converter GUI without a console window
$script   = Join-Path $PSScriptRoot "Convert Word To PDF GUI.ps1"
$desktop  = [Environment]::GetFolderPath('Desktop')
$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path $desktop "Books to PDF.lnk"))

$shortcut.TargetPath       = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments        = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$script`""
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.IconLocation     = Join-Path $PSScriptRoot "Books to PDF.ico"
$shortcut.Description      = "Convert The Guide Word documents to PDF (progress window)"
$shortcut.Save()

Write-Host "Shortcut created: $(Join-Path $desktop 'Books to PDF.lnk')"
