param(
    [string]$TargetPptx = (Join-Path (Split-Path $PSScriptRoot -Parent) "Zephyr_RTOS_Embedded_Systems.pptx"),
    [string]$OutPdf = (Join-Path (Split-Path $PSScriptRoot -Parent) "Zephyr_RTOS_Embedded_Systems.pdf"),
    [switch]$ExportPdf,
    [switch]$ExportPngs
)
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path $PSScriptRoot -Parent
$qaDir = Join-Path ([System.IO.Path]::GetTempPath()) "zephyr_embedded_qa_slides"

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($TargetPptx, $true, $false, $false)

if ($ExportPngs) {
    New-Item -ItemType Directory -Force -Path $qaDir | Out-Null
    Remove-Item (Join-Path $qaDir "Slide*.PNG") -ErrorAction SilentlyContinue
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $pres.Slides.Item($i).Export((Join-Path $qaDir "Slide$i.PNG"), "PNG", 1600, 900)
    }
    Write-Output "EXPORTED_PNGS: $qaDir ($($pres.Slides.Count) slides)"
}

if ($ExportPdf) {
    $pres.SaveAs($OutPdf, 32)
    Write-Output "EXPORTED_PDF: $OutPdf ($((Get-Item $OutPdf).Length) bytes)"
}

$pres.Close()
if ($pp.Presentations.Count -eq 0) {
    $pp.Quit()
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null
