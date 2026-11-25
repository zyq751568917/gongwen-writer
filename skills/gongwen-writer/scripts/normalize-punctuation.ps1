<#
.SYNOPSIS
  把中文语境里的半角标点规范成中文全角标点，输出修改清单。
.DESCRIPTION
  只改"紧邻汉字/中文标点"的半角标点；数字与西文串内部（3.14、1,000、10:30、A-1）不动；
  三级标题序数 "1." 的点号保留半角（国标 GB/T 9704-2012 第7.3.3条）。
.PARAMETER Path
  待处理 docx。
.PARAMETER InPlace
  直接改原文件。
.PARAMETER OutputPath
  另存路径（与 -InPlace 二选一）。
.PARAMETER ReportPath
  修改清单（文本）。
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File normalize-punctuation.ps1 -Path 公文.docx -InPlace -ReportPath 标点修改.txt
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [switch]$InPlace,
  [string]$OutputPath,
  [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $Path)) { Write-Error ("文件不存在: " + $Path); exit 2 }
$full = (Resolve-Path -LiteralPath $Path).Path
if (-not $InPlace -and -not $OutputPath) { Write-Error "请指定 -InPlace 或 -OutputPath"; exit 2 }

$halfMap = @{
  ',' = [string][char]0xFF0C; '.' = [string][char]0x3002; ';' = [string][char]0xFF1B
  ':' = [string][char]0xFF1A; '?' = [string][char]0xFF1F; '!' = [string][char]0xFF01
  '(' = [string][char]0xFF08; ')' = [string][char]0xFF09
}
$cjkRx = [regex]'[一-鿿　-〿＀-￯]'
$alnumRx = [regex]'[A-Za-z0-9]'

$word = $null; $doc = $null
$changes = New-Object System.Collections.ArrayList
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($full, $false, ($false -eq $InPlace), $false)

  $openDouble = $true
  $openSingle = $true
  $paraIndex = 0
  foreach ($p in $doc.Paragraphs) {
    $paraIndex++
    $raw = $p.Range.Text
    $t = $raw.Replace([string][char]13, '').Replace([string][char]7, '')
    if ($t.Length -eq 0) { continue }
    $base = $p.Range.Start
    $skipDot = $t -match '^\s*\d+[\.、]'
    $dotSeen = $false
    for ($i = 0; $i -lt $t.Length; $i++) {
      $ch = [string]$t[$i]
      $prev = ''; $next = ''
      if ($i -gt 0) { $prev = $t[$i - 1] }
      if ($i -lt $t.Length - 1) { $next = $t[$i + 1] }
      $prevAlnum = $prev.Length -gt 0 -and $alnumRx.IsMatch($prev)
      $nextAlnum = $next.Length -gt 0 -and $alnumRx.IsMatch($next)

      if ($ch -eq '"') {
        if ($cjkRx.IsMatch($prev) -or $cjkRx.IsMatch($next)) {
          $new = if ($openDouble) { [string][char]0x201C } else { [string][char]0x201D }
          $openDouble = -not $openDouble
          [void]$changes.Add([PSCustomObject]@{ para = $paraIndex; pos = $base + $i; from = '"'; to = $new })
        }
        continue
      }
      if ($ch -eq "'") {
        if ($cjkRx.IsMatch($prev) -or $cjkRx.IsMatch($next)) {
          $new = if ($openSingle) { [string][char]0x2018 } else { [string][char]0x2019 }
          $openSingle = -not $openSingle
          [void]$changes.Add([PSCustomObject]@{ para = $paraIndex; pos = $base + $i; from = "'"; to = $new })
        }
        continue
      }
      if (-not $halfMap.ContainsKey($ch)) { continue }
      if ($prevAlnum -and $nextAlnum) { continue }
      if ($ch -eq '.' -and $skipDot -and -not $dotSeen) { $dotSeen = $true; continue }
      if ($cjkRx.IsMatch($prev) -or $cjkRx.IsMatch($next)) {
        [void]$changes.Add([PSCustomObject]@{ para = $paraIndex; pos = $base + $i; from = $ch; to = $halfMap[$ch] })
      }
    }
  }

  for ($i = $changes.Count - 1; $i -ge 0; $i--) {
    $c = $changes[$i]
    $r = $doc.Range($c.pos, $c.pos + 1)
    $r.Text = $c.to
  }

  if ($InPlace) { $doc.Save() }
  else {
    $outFull = [System.IO.Path]::GetFullPath($OutputPath)
    if (-not $outFull.ToLowerInvariant().EndsWith('.docx')) { $outFull = $outFull + '.docx' }
    $fmt = 16
    $doc.SaveAs([ref]$outFull, [ref]$fmt)
  }
  $doc.Close(0); $doc = $null
}
catch {
  Write-Error ("标点处理失败: " + $_.Exception.Message)
  exit 3
}
finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($word) { try { $word.Quit() } catch {} }
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("标点修改 " + $changes.Count + " 处")
foreach ($c in $changes) { [void]$sb.AppendLine("第 " + $c.para + " 段: " + $c.from + " -> " + $c.to) }
if ($ReportPath) {
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($ReportPath), $sb.ToString(), $utf8bom)
}
Write-Output $sb.ToString()
exit 0
