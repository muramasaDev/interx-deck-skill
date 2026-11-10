# 장마다 PNG, -Pdf 면 PDF 도. DRM 때문에 사본만 연다
# 사용: render.ps1 -Pptx <파일> -Out <폴더> [-Slides 1,3] [-Pdf <파일>]
param([Parameter(Mandatory)][string]$Pptx, [Parameter(Mandatory)][string]$Out, [string]$Slides, [string]$Pdf)
$ErrorActionPreference = 'Stop'
$src = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path   # COM 내보내기는 절대 경로만 받음
$id = [guid]::NewGuid().ToString('N')
$tmp = Join-Path $env:TEMP "ixrender_$id.pptx"
$tmpPdf = Join-Path $env:TEMP "ixrender_$id.pdf"
Copy-Item $src $tmp
$pp = New-Object -ComObject PowerPoint.Application
try {
    $p = $pp.Presentations.Open($tmp, $true, $false, $false)
    $n = $p.Slides.Count
    $list = if ($Slides) { $Slides -split ',' | ForEach-Object { [int]$_ } } else { 1..$n }
    foreach ($i in $list) { $p.Slides($i).Export((Join-Path $Out "s$i.png"), "PNG", 1600, 900) }
    if ($Pdf) {
        $p.SaveAs($tmpPdf, 32)   # 32 = ppSaveAsPDF
        # Office 가 쓴 파일은 DRM 이 곧 암호화: .NET 으로 다시 씀
        [IO.File]::WriteAllBytes([IO.Path]::GetFullPath($Pdf), [IO.File]::ReadAllBytes($tmpPdf))
    }
    $p.Close()
} finally {
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }   # 사용자가 연 창은 닫지 않음
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    Start-Sleep -Milliseconds 500
    Remove-Item $tmp, $tmpPdf -Force -ErrorAction SilentlyContinue
}
Get-ChildItem $Out -Filter s*.png | ForEach-Object { $_.FullName }
