<#
.SYNOPSIS
  抽取 Word / Excel / PPT / PDF 材料文字。
.DESCRIPTION
  docx / doc / pdf  -> Word COM；xlsx -> Excel COM；pptx -> PowerPoint COM。
  PDF 走 Word 重排，无文本层时明确提示"疑似扫描件，需 OCR"，不猜测内容。
.PARAMETER Path
  材料文件路径。
.PARAMETER OutFile
  纯文本输出路径（UTF-8）。省略则打到标准输出。
.PARAMETER JsonPath
  JSON 输出路径（结构化：type / index / text）。
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File extract-text.ps1 -Path "材料.docx" -OutFile "材料.txt"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$OutFile,
  [string]$JsonPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $Path)) { Write-Error ("文件不存在: " + $Path); exit 2 }
$full = (Resolve-Path -LiteralPath $Path).Path
$ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()

$items = New-Object System.Collections.ArrayList
$notes = New-Object System.Collections.ArrayList

function Add-Item {
  param([string]$Type, [int]$Index, [string]$Text)
  [void]$items.Add([PSCustomObject]@{ type = $Type; index = $Index; text = $Text })
}

switch ($ext) {
  { $_ -in @('.docx', '.doc', '.pdf') } {
    $word = $null; $doc = $null
    try {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $word.DisplayAlerts = 0
      # ConfirmConversions=false 让 PDF 静默重排，ReadOnly=true
      $doc = $word.Documents.Open($full, $false, $true, $false)
      $i = 0
      foreach ($p in $doc.Paragraphs) {
        $i++
        $t = $p.Range.Text
        $t = $t.Replace([string][char]13, '').Replace([string][char]7, '')
        Add-Item 'paragraph' $i $t
      }
    }
    catch {
      Write-Error ("Word 打开失败: " + $_.Exception.Message); exit 4
    }
    finally {
      if ($doc) { try { $doc.Close(0) } catch {} }
      if ($word) { try { $word.Quit() } catch {} }
    }
    $chars = ($items | ForEach-Object { $_.text }) -join ''
    if ($ext -eq '.pdf' -and $chars.Trim().Length -lt 10) {
      [void]$notes.Add("PDF 无文本层（疑似扫描件），本机无 OCR，需人工识别后再处理。")
    }
  }

  '.xlsx' {
    $excel = $null; $wb = $null
    try {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
      $wb = $excel.Workbooks.Open($full, $false, $true)
      $si = 0
      foreach ($ws in $wb.Worksheets) {
        $si++
        Add-Item 'sheet' $si ("=== 工作表: " + $ws.Name + " ===")
        $range = $ws.UsedRange
        if ($null -eq $range) { continue }
        $vals = $range.Value2
        $rowCount = $range.Rows.Count
        $colCount = $range.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
          $cells = @()
          for ($c = 1; $c -le $colCount; $c++) {
            $v = $null
            if ($rowCount -eq 1 -and $colCount -eq 1) { $v = $vals }
            else { $v = $vals.GetValue($r, $c) }
            if ($null -ne $v) { $cells += ([string]$v) } else { $cells += '' }
          }
          $line = ($cells -join ' ').TrimEnd()
          if ($line.Trim().Length -gt 0) { Add-Item 'row' $r $line }
        }
      }
    }
    catch {
      Write-Error ("Excel 打开失败: " + $_.Exception.Message); exit 4
    }
    finally {
      if ($wb) { try { $wb.Close($false) } catch {} }
      if ($excel) { try { $excel.Quit() } catch {} }
    }
  }

  '.pptx' {
    $ppt = $null; $pres = $null
    try {
      $ppt = New-Object -ComObject PowerPoint.Application
      $pres = $ppt.Presentations.Open($full, $true, $false, $false)
      $si = 0
      foreach ($slide in $pres.Slides) {
        $si++
        Add-Item 'slide' $si ("=== 幻灯片 " + $si + " ===")
        $idx = 0
        foreach ($shape in $slide.Shapes) {
          $idx++
          try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
              $t = $shape.TextFrame.TextRange.Text.Replace([string][char]13, [string][char]10).Trim()
              if ($t.Length -gt 0) { Add-Item 'shape' $idx $t }
            }
          }
          catch { }
        }
      }
    }
    catch {
      Write-Error ("PowerPoint 打开失败: " + $_.Exception.Message); exit 4
    }
    finally {
      if ($pres) { try { $pres.Close() } catch {} }
      if ($ppt) { try { $ppt.Quit() } catch {} }
    }
  }

  { $_ -in @('.txt', '.md', '.csv') } {
    $raw = Get-Content -LiteralPath $full -Raw -Encoding UTF8
    $i = 0
    foreach ($line in ($raw -split [char]10)) {
      $i++
      if ($line.Trim().Length -gt 0) { Add-Item 'line' $i $line.TrimEnd() }
    }
  }

  default {
    Write-Error ("不支持的文件类型: " + $ext + "（支持 .docx/.doc/.pdf/.xlsx/.pptx/.txt/.md/.csv）")
    exit 2
  }
}

$plain = New-Object System.Text.StringBuilder
if ($notes.Count -gt 0) {
  foreach ($n in $notes) { [void]$plain.AppendLine("[注意] " + $n) }
}
foreach ($it in $items) { [void]$plain.AppendLine($it.text) }
$plainText = $plain.ToString()

if ($OutFile) {
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutFile), $plainText, $utf8bom)
  Write-Output ("已写出: " + [System.IO.Path]::GetFullPath($OutFile))
}
else { Write-Output $plainText }

if ($JsonPath) {
  $payload = [PSCustomObject]@{
    source  = $full
    notes   = @($notes)
    items   = @($items)
  }
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  $utf8noBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($JsonPath), ($payload | ConvertTo-Json -Depth 6), $utf8noBom)
  Write-Output ("已写出: " + [System.IO.Path]::GetFullPath($JsonPath))
}

if ($items.Count -eq 0) {
  Write-Output "[注意] 未抽取到任何文本"
  exit 1
}
exit 0
