# Creates a "Books to PDF" shortcut on the desktop that runs the converter in a console window
$script   = Join-Path $PSScriptRoot "Convert Word To PDF.ps1"
$desktop  = [Environment]::GetFolderPath('Desktop')
$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path $desktop "Books to PDF.lnk"))

$shortcut.TargetPath       = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments        = "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.IconLocation     = "$env:SystemRoot\System32\imageres.dll,2"
$shortcut.Description      = "Convert The Guide Word documents to PDF"
$shortcut.Save()

Write-Host "Shortcut created: $(Join-Path $desktop 'Books to PDF.lnk')"
