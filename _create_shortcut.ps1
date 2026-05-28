param([string]$AppDir)

$AppDir    = $AppDir.TrimEnd('\')
$LaunchVbs = Join-Path $AppDir "launch.vbs"
$IconPath  = Join-Path $AppDir "static\favicon.ico"
$Desktop   = [Environment]::GetFolderPath("Desktop")
$Shortcut  = Join-Path $Desktop "PM CV Formatter.lnk"

try {
    $WshShell             = New-Object -ComObject WScript.Shell
    $lnk                  = $WshShell.CreateShortcut($Shortcut)
    $lnk.TargetPath       = "wscript.exe"
    $lnk.Arguments        = """$LaunchVbs"""
    $lnk.WorkingDirectory = $AppDir
    $lnk.Description      = "PM CV Formatter"
    $lnk.WindowStyle      = 1
    if (Test-Path $IconPath) { $lnk.IconLocation = "$IconPath,0" }
    $lnk.Save()
} catch {
    exit 1
}

exit 0
