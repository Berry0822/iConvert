param(
  [string]$Pyw    = $env:ICONV_PYW,
  [string]$AppDir = $env:ICONV_APPDIR
)
$script = Join-Path $AppDir 'file_converter.py'
$icon   = Join-Path $AppDir 'icon.ico'
$shell  = New-Object -ComObject WScript.Shell

function New-Lnk([string]$linkPath) {
  $s = $shell.CreateShortcut($linkPath)
  $s.TargetPath       = $Pyw
  $s.Arguments        = '"' + $script + '"'
  $s.WorkingDirectory = $AppDir
  $s.IconLocation     = $icon
  $s.Description       = 'iConvert - Local File Converter'
  $s.Save()
}

New-Lnk (Join-Path ([Environment]::GetFolderPath('Programs')) 'iConvert.lnk')
New-Lnk (Join-Path ([Environment]::GetFolderPath('Desktop'))  'iConvert.lnk')
Write-Host 'Shortcuts created.'
