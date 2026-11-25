<#
.SYNOPSIS
  校验 Word 公文是否符合 GB/T 9704-2012 版式与字体规范。
.DESCRIPTION
  检查项：页面边距与文档网格、逐段中文/西文字体与字号、首行缩进、行距、
  页码（宋体 4 号一字线、奇偶分侧）、中文语境半角标点、正文是否落在首页。
  只读不改。退出码 0 = 无 FAIL，1 = 有 FAIL。
.PARAMETER Path
  待校验的 docx。
.PARAMETER JsonPath
  报告 JSON 输出路径（省略打到标准输出）。
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File verify-gongwen.ps1 -Path 请示.docx
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$JsonPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$FONT_TITLE = '方正小标宋_GBK'
$FONT_H1    = '方正黑体_GBK'
$FONT_H2    = '方正楷体_GBK'
$FONT_BODY  = '方正仿宋_GBK'
$FONT_WEST  = 'Times New Roman'
$FONT_PAGE  = '宋体'

$Expect = @{
  redHeader = @($FONT_TITLE, 22); title = @($FONT_TITLE, 22); attachmentTitle = @($FONT_TITLE, 22)
  h1 = @($FONT_H1, 16); h2 = @($FONT_H2, 16)
  body = @($FONT_BODY, 16); h3 = @($FONT_BODY, 16); h4 = @($FONT_BODY, 16)
  mainRecipient = @($FONT_BODY, 16); docNumber = @($FONT_BODY, 16); signer = @($FONT_BODY, 16)
  attachmentNote = @($FONT_BODY, 16); issuer = @($FONT_BODY, 16); date = @($FONT_BODY, 16)
  note = @($FONT_BODY, 16); attachmentMarker = @($FONT_H1, 16)
  cc = @($FONT_BODY, 14); printer = @($FONT_BODY, 14)
}

function Detect-Role {
  param([string]$t, [int]$idx)
  $s = $t.Trim()
  if ($s.Length -eq 0) { return 'empty' }
  if ($idx -eq 1 -and $s -match '文件$') { return 'redHeader' }
  if ($s -match '^签发人') { return 'signer' }
  if ($s -match '^(抄送|抄报)') { return 'cc' }
  if ($s -match '印发$') { return 'printer' }
  if ($s -match '^附件[:：]') { return 'attachmentNote' }
  if ($s -match '^附件[一二三四五六七八九十\d]*$') { return 'attachmentMarker' }
  if ($s -match '^[一二三四五六七八九十百]+、') { return 'h1' }
  if ($s -match '^（[一二三四五六七八九十]+）') { return 'h2' }
  if ($s -match '^\d+[\.、]\s*') { return 'h3' }
  if ($s -match '^[（(]\d+[）)]') { return 'h4' }
  if ($s -match '〔\d{4}〕.+号$') { return 'docNumber' }
  if ($s -match '^\d{4}年\d{1,2}月\d{1,2}日$') { return 'date' }
  if ($idx -lt 12 -and $s.Length -le 40 -and $s -notmatch '[。；：]$' -and $s -match '(请示|报告|通知|批复|函|通报|意见|决定|议案|纪要|公告|通告|命令)$') { return 'title' }
  if ($s -match '[:：]$' -and $s.Length -le 30) { return 'mainRecipient' }
  return 'body'
}

$halfMap = @{
  ',' = [string][char]0xFF0C; '.' = [string][char]0x3002; ';' = [string][char]0xFF1B
  ':' = [string][char]0xFF1A; '?' = [string][char]0xFF1F; '!' = [string][char]0xFF01
  '(' = [string][char]0xFF08; ')' = [string][char]0xFF09
}
$cjkRx = [regex]'[一-鿿　-〿＀-￯]'
$alnumRx = [regex]'[A-Za-z0-9]'

function Find-HalfPunct {
  param([string]$text)
  $hits = @()
  for ($i = 0; $i -lt $text.Length; $i++) {
    $ch = [string]$text[$i]
    if (-not $halfMap.ContainsKey($ch)) { continue }
    $prev = ''; $next = ''
    if ($i -gt 0) { $prev = $text[$i - 1] }
    if ($i -lt $text.Length - 1) { $next = $text[$i + 1] }
    $prevAlnum = $prev.Length -gt 0 -and $alnumRx.IsMatch($prev)
    $nextAlnum = $next.Length -gt 0 -and $alnumRx.IsMatch($next)
    if ($prevAlnum -and $nextAlnum) { continue }
    if ($cjkRx.IsMatch($prev) -or $cjkRx.IsMatch($next)) {
      $hits += [PSCustomObject]@{ index = $i; char = $ch; replace = $halfMap[$ch] }
    }
  }
  return $hits
}

if (-not (Test-Path -LiteralPath $Path)) { Write-Error ("文件不存在: " + $Path); exit 2 }
$full = (Resolve-Path -LiteralPath $Path).Path

$word = $null; $doc = $null
$fails = New-Object System.Collections.ArrayList
$warns = New-Object System.Collections.ArrayList
$paras = New-Object System.Collections.ArrayList
$punct = New-Object System.Collections.ArrayList

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($full, $false, $true, $false)

  $ps = $doc.PageSetup
  $pageInfo = [PSCustomObject]@{
    paperSize = $ps.PaperSize
    marginTopMm = [math]::Round($ps.TopMargin / 72 * 25.4, 2)
    marginBottomMm = [math]::Round($ps.BottomMargin / 72 * 25.4, 2)
    marginLeftMm = [math]::Round($ps.LeftMargin / 72 * 25.4, 2)
    marginRightMm = [math]::Round($ps.RightMargin / 72 * 25.4, 2)
    layoutMode = $ps.LayoutMode
    gridLines = $ps.LinesPage
    gridChars = $ps.CharsLine
    oddEvenFooter = $ps.OddAndEvenPagesHeaderFooter
  }
  $expectPage = @{ marginTopMm = 37; marginBottomMm = 35; marginLeftMm = 28; marginRightMm = 26 }
  foreach ($k in $expectPage.Keys) {
    if ([math]::Abs($pageInfo.$k - $expectPage[$k]) -gt 0.6) {
      [void]$fails.Add(("页面 " + $k + " = " + $pageInfo.$k + "mm，期望 " + $expectPage[$k] + "mm"))
    }
  }
  if ($pageInfo.layoutMode -ne 1) { [void]$fails.Add("未启用文档网格（LayoutMode 应为 1）") }
  if ($pageInfo.gridLines -ne 22) { [void]$fails.Add("每页行数 = " + $pageInfo.gridLines + "，期望 22") }
  if ($pageInfo.gridChars -ne 28) { [void]$fails.Add("每行字数 = " + $pageInfo.gridChars + "，期望 28") }

  $roles = @()
  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    $t = $p.Range.Text.Replace([string][char]13, '').Replace([string][char]7, '')
    $roles += (Detect-Role $t $i)
  }
  for ($i = 0; $i -lt $roles.Count; $i++) {
    if ($roles[$i] -eq 'body' -and ($i + 1) -lt $roles.Count -and $roles[$i + 1] -eq 'date') { $roles[$i] = 'issuer' }
  }

  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    $t = $p.Range.Text.Replace([string][char]13, '').Replace([string][char]7, '')
    $role = $roles[$i - 1]
    $f = $p.Range.Font
    $pf = $p.Format
    $entry = [PSCustomObject]@{
      index = $i; role = $role; text = $t
      fontFarEast = $p.Range.Characters.Item(1).Font.NameFarEast
      fontAscii = $p.Range.Characters.Item(1).Font.NameAscii
      sizePt = $p.Range.Characters.Item(1).Font.Size
      align = $pf.Alignment; indentFirst = $pf.CharacterUnitFirstLineIndent
      indentLeft = $pf.CharacterUnitLeftIndent; indentRight = $pf.CharacterUnitRightIndent
      lineSpacingRule = $pf.LineSpacingRule; lineSpacing = $pf.LineSpacing
      page = $p.Range.Information(3)
    }
    [void]$paras.Add($entry)

    if ($Expect.ContainsKey($role)) {
      $exp = $Expect[$role]
      $allowed = @($exp[0])
      if ($role -eq 'signer') { $allowed = @($FONT_BODY, $FONT_H2) }
      $c1 = $p.Range.Characters.Item(1)
      $c1Font = $c1.Font.NameFarEast
      $c1Size = $c1.Font.Size
      if ($c1Font -ne $exp[0]) {
        [void]$fails.Add(("第 " + $i + " 段(" + $role + ") 中文字体 " + $c1Font + "，期望 " + $exp[0]))
      }
      if ([math]::Abs($c1Size - $exp[1]) -gt 0.3) {
        [void]$fails.Add(("第 " + $i + " 段(" + $role + ") 字号 " + $c1Size + "pt，期望 " + $exp[1] + "pt"))
      }
      if ($f.NameFarEast -eq '') {
        $seen = @{}
        $scan = [math]::Min($p.Range.Characters.Count, 200)
        for ($ci = 1; $ci -le $scan; $ci++) {
          $cf = $p.Range.Characters.Item($ci).Font.NameFarEast
          if ($cf -and -not $seen.ContainsKey($cf)) { $seen[$cf] = $true }
        }
        foreach ($cf in $seen.Keys) {
          if ($allowed -notcontains $cf) {
            [void]$fails.Add(("第 " + $i + " 段(" + $role + ") 出现非预期中文字体 " + $cf + "（允许: " + ($allowed -join '/') + "）"))
          }
        }
      }
      $asciiName = $f.NameAscii
      if (-not $asciiName) { $asciiName = $c1.Font.NameAscii }
      if ($asciiName -ne $FONT_WEST) {
        [void]$fails.Add(("第 " + $i + " 段(" + $role + ") 西文字体 " + $asciiName + "，期望 " + $FONT_WEST))
      }
      if ($role -in @('body', 'h1', 'h2', 'h3', 'h4') -and [math]::Abs($pf.CharacterUnitFirstLineIndent - 2) -gt 0.2) {
        [void]$fails.Add(("第 " + $i + " 段(" + $role + ") 首行缩进 " + $pf.CharacterUnitFirstLineIndent + " 字符，期望 2"))
      }
      if ($role -eq 'attachmentNote' -and [math]::Abs($pf.CharacterUnitLeftIndent - 2) -gt 0.2) {
        [void]$fails.Add(("第 " + $i + " 段 附件说明左缩进 " + $pf.CharacterUnitLeftIndent + " 字符，期望 2"))
      }
      if ($role -eq 'date' -and [math]::Abs($pf.CharacterUnitRightIndent - 4) -gt 0.2) {
        [void]$fails.Add(("第 " + $i + " 段 成文日期右缩进 " + $pf.CharacterUnitRightIndent + " 字符，期望 4"))
      }
    }
    foreach ($h in (Find-HalfPunct $t)) {
      [void]$punct.Add([PSCustomObject]@{ paragraph = $i; char = $h.char; replace = $h.replace; context = $t.Substring([math]::Max(0, $h.index - 8), [math]::Min(18, $t.Length - [math]::Max(0, $h.index - 8))) })
    }
  }
  if ($punct.Count -gt 0) {
    [void]$fails.Add("中文语境半角标点 " + $punct.Count + " 处（详见标点清单）")
  }

  $firstBody = $paras | Where-Object { $_.role -in @('body', 'h1') } | Select-Object -First 1
  if ($firstBody -and $firstBody.page -ne 1) {
    [void]$fails.Add("正文首段落在第 " + $firstBody.page + " 页，国标要求首页显示正文")
  }
  elseif (-not $firstBody) { [void]$warns.Add("未发现正文段落") }

  $sec = $doc.Sections.Item(1)
  $footOdd = $sec.Footers.Item(1).Range
  $footEven = $sec.Footers.Item(3).Range
  $fn = [PSCustomObject]@{
    oddText = $footOdd.Text.Trim(); evenText = $footEven.Text.Trim()
    oddFont = $footOdd.Font.NameFarEast; oddSize = $footOdd.Font.Size; oddAlign = $footOdd.Paragraphs.Item(1).Format.Alignment
    evenFont = $footEven.Font.NameFarEast; evenSize = $footEven.Font.Size; evenAlign = $footEven.Paragraphs.Item(1).Format.Alignment
    oddHasField = ($footOdd.Fields.Count -gt 0); evenHasField = ($footEven.Fields.Count -gt 0)
  }
  if ($fn.oddText.Length -eq 0 -and $fn.evenText.Length -eq 0) { [void]$fails.Add("缺少页码") }
  else {
    if ($fn.oddFont -ne $FONT_PAGE) { [void]$fails.Add("页码字体 " + $fn.oddFont + "，期望 " + $FONT_PAGE) }
    if ([math]::Abs($fn.oddSize - 14) -gt 0.3) { [void]$fails.Add("页码字号 " + $fn.oddSize + "pt，期望 14pt（4号）") }
    if (-not $fn.oddHasField) { [void]$fails.Add("页脚缺 PAGE 域") }
    if ($fn.oddText.IndexOf([string][char]0x2014) -lt 0) { [void]$fails.Add("页码缺一字线") }
    if ($fn.oddAlign -ne 2) { [void]$fails.Add("单页页码应居右") }
    if ($fn.evenAlign -ne 0) { [void]$fails.Add("双页页码应居左") }
  }

  $pages = $doc.ComputeStatistics(2)
  $doc.Close(0); $doc = $null
}
catch {
  Write-Error ("校验失败（第 " + $_.InvocationInfo.ScriptLineNumber + " 行）: " + $_.Exception.Message)
  exit 3
}
finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($word) { try { $word.Quit() } catch {} }
}

$report = [PSCustomObject]@{
  file = $full
  pages = $pages
  pass = ($fails.Count -eq 0)
  fails = @($fails)
  warns = @($warns)
  page = $pageInfo
  pageNumber = $fn
  punctuation = @($punct)
  paragraphs = @($paras)
}

$json = $report | ConvertTo-Json -Depth 6
if ($JsonPath) {
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  $utf8noBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($JsonPath), $json, $utf8noBom)
}
else { Write-Output $json }

Write-Output ("共 " + $paras.Count + " 段，" + $pages + " 页；FAIL " + $fails.Count + " 项")
foreach ($f in $fails) { Write-Output ("[FAIL] " + $f) }
foreach ($w in $warns) { Write-Output ("[WARN] " + $w) }
if ($fails.Count -eq 0) { Write-Output "[OK] 版式合规" }
exit $(if ($fails.Count -eq 0) { 0 } else { 1 })
