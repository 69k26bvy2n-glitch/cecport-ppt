# 安装技能内置 MiSans 字体到当前用户（Windows，无需管理员）
$src = Join-Path $PSScriptRoot "..\assets\fonts"
$dest = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"

if (-not (Test-Path $src)) {
  Write-Error "未找到字体目录：$src"
  exit 1
}

New-Item -ItemType Directory -Force -Path $dest | Out-Null
Get-ChildItem $src -Filter *.ttf | ForEach-Object {
  Copy-Item $_.FullName $dest -Force
  $label = "MiSans {0} (TrueType)" -f ($_.BaseName -replace '^MiSans-', '')
  $target = Join-Path $dest $_.Name
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" `
    -Name $label -Value $target -PropertyType String -Force | Out-Null
}
Write-Output "已安装到 $dest，重启 PowerPoint/WPS 后生效。"
