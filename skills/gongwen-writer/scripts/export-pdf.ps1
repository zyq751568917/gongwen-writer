<#
.SYNOPSIS
  把公文 docx 导出为 PDF（可选预览件）。
.PARAMETER Path
  docx 路径。
.PARAMETER OutputPath
  PDF 输出路径，默认与源文件同名。
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File export-pdf.ps1 -Path 公文.docx -OutputPath 公文.pdf
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $Path)) { Write-Error ("文件不存在: " + $Path); exit 2 }
$full = (Resolve-Path -LiteralPath $Path).Path
if (-not $OutputPath) {
  $OutputPath = [System.IO.Path]::ChangeExtension($full, '.pdf')
}
$outFull = [System.IO.Path]::GetFullPath($OutputPath)

$word = $null; $doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($full, $false, $true, $false)
  $doc.ExportAsFixedFormat($outFull, 17)
  $doc.Close(0); $doc = $null
  Write-Output ("已导出: " + $outFull)
  exit 0
}
catch {
  Write-Error ("导出失败: " + $_.Exception.Message)
  exit 3
}
finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($word) { try { $word.Quit() } catch {} }
}
