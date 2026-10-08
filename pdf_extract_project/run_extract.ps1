param(
    [string]$PdfPath = "E:\news\newschina170 no.pdf",
    [string]$OutputName = "newschina_output"
)

$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $projectDir "extract_news.py"
$outputDir = Join-Path $projectDir "output"
$jsonPath = Join-Path $outputDir ($OutputName + ".json")
$docxPath = Join-Path $outputDir ($OutputName + ".docx")

if (-not (Test-Path $PdfPath)) {
    Write-Error "PDF file not found: $PdfPath"
}

if (-not (Test-Path $scriptPath)) {
    Write-Error "extract_news.py not found: $scriptPath"
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

Write-Host "PDF:" $PdfPath
Write-Host "JSON:" $jsonPath
Write-Host "DOCX:" $docxPath

python $scriptPath --pdf $PdfPath --json $jsonPath --docx $docxPath

Write-Host ""
Write-Host "Finished."
