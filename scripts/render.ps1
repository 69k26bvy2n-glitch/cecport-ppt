# render.ps1 — Windows 渲染：PowerPoint/WPS COM 直接导出逐页 PNG（零第三方依赖）
# 用法: powershell -ExecutionPolicy Bypass -File render.ps1 <输入.pptx> <输出目录> [宽度px]
# 输出统一为: 输出目录\slides\slides.001.png ...（与 macOS render.sh 同名规范）
# 说明：Windows 版比 Mac 简单——COM 一步出 PNG，无需 PDF 中转，无沙盒授权弹窗。
param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [int]$Width = 1600
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $Path)) { Write-Host "找不到文件: $Path"; exit 1 }
$absIn  = (Resolve-Path -LiteralPath $Path).Path
if ([System.IO.Path]::IsPathRooted($OutDir)) { $absOut = $OutDir } else { $absOut = Join-Path (Get-Location) $OutDir }
New-Item -ItemType Directory -Force -Path "$absOut\slides" | Out-Null

# 微信/浏览器下载的 pptx 带"网络下载标记"会触发受保护视图导致 COM 打不开，先解除
Unblock-File -LiteralPath $absIn -ErrorAction SilentlyContinue

# 优先微软 PowerPoint；没装则回退 WPS 演示（国内常见，COM 接口同名兼容）
$app = $null
foreach ($progId in "PowerPoint.Application","KWPP.Application") {
    try { $app = New-Object -ComObject $progId; break } catch {}
}
if (-not $app) {
    Write-Host "本机没有 PowerPoint 或 WPS，无法自动渲染。请手工导出 PDF/图片后检查。"
    exit 2
}

$pres = $null
try {
    # Open(FileName, ReadOnly, Untitled, WithWindow) —— 只读、不建窗口，后台导出
    $pres = $app.Presentations.Open($absIn, $true, $true, $false)
    $ratio  = $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth
    $height = [int]($Width * $ratio)
    $n = $pres.Slides.Count
    for ($i = 1; $i -le $n; $i++) {
        # 逐页导出并显式命名：整体导出在中文版 Office 里文件名是「幻灯片1.PNG」，不规范
        $pres.Slides.Item($i).Export("$absOut\slides\slides.{0:d3}.png" -f $i, "PNG", $Width, $height)
    }
    Write-Host "[render] 完成: $absOut\slides\ （共 $n 页）"
} catch {
    Write-Host "[render] 导出失败: $($_.Exception.Message)"
    Write-Host "常见原因：① 文件本身弹了修复/受保护视图框——先手工用 PowerPoint 打开确认一次；② Office 尚未完成首次启动向导。"
    exit 3
} finally {
    if ($pres) { try { $pres.Close() } catch {} }
    $app.Quit()
}
