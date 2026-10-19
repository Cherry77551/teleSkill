<#
.SYNOPSIS
  Render a .docx to PDF and export page images (JPEG) for visual check.
.DESCRIPTION
  默认用 LibreOffice (soffice) 转 PDF，再用 Poppler (pdftoppm) 出图。
  加 -UseWord 改用本机 Word（COM）导出 PDF——与 Word 内显示一致，可避免
  LibreOffice 缺数学字体导致的公式假失败；需要装 Microsoft Word。

  Usage:
    .\render_check.ps1 -Docx out.docx                              # 默认抽样：只渲染第 1 页
    .\render_check.ps1 -Docx out.docx -FirstPage 7 -LastPage 9     # 抽样：只渲染 7-9 页
    .\render_check.ps1 -Docx out.docx -All                          # 全量渲染（仅交付前一次性可选）
    .\render_check.ps1 -Docx out.docx -FirstPage 1 -LastPage 3 -Dpi 150
    .\render_check.ps1 -Docx out.docx -UseWord -FirstPage 7 -LastPage 7 # 用 Word 渲染公式页

  Performance: 渲染是长流程的大头，默认只抽渲染公式页、目录页、首页/尾页，
  用 -FirstPage / -LastPage 指定范围；全页渲染（-All）仅最终交付前一次性可选。
  中间产物（PDF + 页面图）一律写到 <docx 目录>\.temp\，不进交付目录。
#>
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [int]$FirstPage = 1,
    [int]$LastPage = 0,
    [switch]$All,
    [int]$Dpi = 100,
    [switch]$UseWord,
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command pdftoppm -ErrorAction SilentlyContinue)) {
    Write-Error "pdftoppm (Poppler) not found. Install Poppler and ensure pdftoppm is in PATH."
    exit 1
}
if (-not $UseWord -and -not (Get-Command soffice -ErrorAction SilentlyContinue)) {
    Write-Error "soffice (LibreOffice) not found. Install LibreOffice, or pass -UseWord. Ensure it is in PATH."
    exit 1
}

$full = (Resolve-Path $Docx).Path
$base = [System.IO.Path]::GetFileNameWithoutExtension($full)
$dir  = Split-Path -Parent $full
if (-not $OutDir) { $OutDir = Join-Path $dir ".temp" }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$pdf  = Join-Path $OutDir "$base.pdf"

if ($UseWord) {
    Write-Host "Converting to PDF via Word COM..."
    $word = $null; $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($full, $false, $true)
        $doc.ExportAsFixedFormat($pdf, 17)   # wdExportFormatPDF = 17
        $doc.Close($false)
    } finally {
        if ($doc)  { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    }
    if (-not (Test-Path $pdf)) {
        Write-Error "Word COM PDF export failed (no output at $pdf)."
        exit 1
    }
} else {
    Write-Host "Converting to PDF via LibreOffice..."
    soffice --headless --convert-to pdf --outdir $OutDir $full
    if ($LASTEXITCODE -ne 0) {
        Write-Error "PDF conversion failed (soffice exit code $LASTEXITCODE)"
        exit 1
    }
}

$prefix = Join-Path $OutDir "${base}-page"
if ($All) {
    Write-Host "Rendering all pages (final optional pass)..."
    pdftoppm -jpeg -r $Dpi $pdf $prefix
} elseif ($LastPage -gt 0) {
    Write-Host "Rendering pages $FirstPage..$LastPage (sampled)..."
    pdftoppm -f $FirstPage -l $LastPage -jpeg -r $Dpi $pdf $prefix
} else {
    Write-Host "Rendering page $FirstPage (sampled, default)..."
    pdftoppm -f $FirstPage -l $FirstPage -jpeg -r $Dpi $pdf $prefix
}
if ($LASTEXITCODE -ne 0) {
    Write-Error "pdftoppm failed (exit code $LASTEXITCODE)"
    exit 1
}

Write-Host "Done: $pdf"
Write-Host "Page images: $OutDir\${base}-page-*.jpg"
