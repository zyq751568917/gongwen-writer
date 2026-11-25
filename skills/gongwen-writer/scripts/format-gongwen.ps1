<#
.SYNOPSIS
  按 GB/T 9704-2012 把 spec.json 排成规范公文 Word（Engine A，Word COM）。
.DESCRIPTION
  - 输入：spec.json（内容 + 段落角色）
  - 输出：合规 .docx，另附排版报告 JSON
  - 字体分槽：中文槽用方正 GBK 四款，西文/数字槽用 Times New Roman，页码用宋体
  - 页面：A4，边距 37/35/28/26 mm，文档网格每页 22 行每行 28 字
.PARAMETER SpecPath
  spec.json 路径。
.PARAMETER OutputPath
  输出 docx 路径。
.PARAMETER ReportPath
  排版报告 JSON 路径（省略打到标准输出）。
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File format-gongwen.ps1 -SpecPath spec.json -OutputPath 请示.docx
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$SpecPath,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [string]$ReportPath,
  [switch]$StrictFont
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$CR = [string][char]13
$VT = [string][char]11

$FONT_TITLE = '方正小标宋_GBK'
$FONT_H1    = '方正黑体_GBK'
$FONT_H2    = '方正楷体_GBK'
$FONT_BODY  = '方正仿宋_GBK'
$FONT_WEST  = 'Times New Roman'
$FONT_PAGE  = '宋体'

# 角色默认值：Font 中文槽 / Size 磅 / Align left|center|right|justify /
# First 首行缩进字符 / Left 左缩进字符 / Right 右缩进字符 / Line 行距磅 / Before 段前空行
$RoleDefaults = @{
  redHeader      = @{ Font = $FONT_TITLE; Size = 22; Align = 'center';  First = 0; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  redLine        = @{ Font = $FONT_BODY;  Size = 6;  Align = 'left';   First = 0; Left = 0; Right = 0; Line = 1;  Before = 0; After = 0; Color = 'auto'; Bold = $false }
  docNumber      = @{ Font = $FONT_BODY;  Size = 16; Align = 'center'; First = 0; Left = 0; Right = 0; Line = 29; Before = 2; After = 0; Color = 'auto'; Bold = $false }
  signer         = @{ Font = $FONT_BODY;  Size = 16; Align = 'right';  First = 0; Left = 0; Right = 1; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  title          = @{ Font = $FONT_TITLE; Size = 22; Align = 'center'; First = 0; Left = 0; Right = 0; Line = 29; Before = 2; After = 0; Color = 'auto'; Bold = $false }
  mainRecipient  = @{ Font = $FONT_BODY;  Size = 16; Align = 'left';   First = 0; Left = 0; Right = 0; Line = 29; Before = 1; After = 0; Color = 'auto'; Bold = $false }
  body           = @{ Font = $FONT_BODY;  Size = 16; Align = 'justify'; First = 2; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  h1             = @{ Font = $FONT_H1;    Size = 16; Align = 'left';   First = 2; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  h2             = @{ Font = $FONT_H2;    Size = 16; Align = 'left';   First = 2; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  h3             = @{ Font = $FONT_BODY;  Size = 16; Align = 'left';   First = 2; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  h4             = @{ Font = $FONT_BODY;  Size = 16; Align = 'left';   First = 2; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  attachmentNote = @{ Font = $FONT_BODY;  Size = 16; Align = 'left';   First = 0; Left = 2; Right = 0; Line = 29; Before = 1; After = 0; Color = 'auto'; Bold = $false }
  issuer         = @{ Font = $FONT_BODY;  Size = 16; Align = 'right';  First = 0; Left = 0; Right = 0; Line = 29; Before = 3; After = 0; Color = 'auto'; Bold = $false }
  date           = @{ Font = $FONT_BODY;  Size = 16; Align = 'right';  First = 0; Left = 0; Right = 4; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  note           = @{ Font = $FONT_BODY;  Size = 16; Align = 'left';   First = 0; Left = 2; Right = 0; Line = 29; Before = 1; After = 0; Color = 'auto'; Bold = $false }
  cc             = @{ Font = $FONT_BODY;  Size = 14; Align = 'left';   First = 0; Left = 1; Right = 1; Line = 21; Before = 1; After = 0; Color = 'auto'; Bold = $false }
  printer        = @{ Font = $FONT_BODY;  Size = 14; Align = 'left';   First = 0; Left = 1; Right = 1; Line = 21; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  attachmentTitle  = @{ Font = $FONT_TITLE; Size = 22; Align = 'center'; First = 0; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
  attachmentMarker = @{ Font = $FONT_H1;    Size = 16; Align = 'left';   First = 0; Left = 0; Right = 0; Line = 29; Before = 0; After = 0; Color = 'auto'; Bold = $false }
}

$AlignMap = @{ left = 0; center = 1; right = 2; justify = 3 }
$ColorMap = @{ auto = -16777216; black = 0; red = 255 }

if (-not (Test-Path -LiteralPath $SpecPath)) { Write-Error ("spec 不存在: " + $SpecPath); exit 2 }
$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$blocks = @($spec.blocks)
if ($blocks.Count -eq 0) { Write-Error "spec.blocks 为空"; exit 2 }

$unknown = @()
for ($i = 0; $i -lt $blocks.Count; $i++) {
  $r = [string]$blocks[$i].role
  if (-not $RoleDefaults.ContainsKey($r)) { $unknown += ("#" + $i + ":" + $r) }
}
if ($unknown.Count -gt 0) { Write-Error ("未知角色: " + ($unknown -join ', ')); exit 2 }

$texts = @()
for ($i = 0; $i -lt $blocks.Count; $i++) {
  $t = [string]$blocks[$i].text
  if ([string]$blocks[$i].role -eq 'redLine') { if ($t.Trim().Length -eq 0) { $t = ' ' } }
  elseif ($t.Length -eq 0) { Write-Error ("block #" + $i + " 缺少 text"); exit 2 }
  $t = $t.Replace([string][char]10, $VT).Replace([string][char]13, $VT)
  $texts += $t
}

$installed = @()
try {
  Add-Type -AssemblyName System.Drawing
  $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }
} catch { }
$regNames = @()
foreach ($regKey in @('HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')) {
  try {
    if (Test-Path $regKey) {
      foreach ($prop in (Get-ItemProperty -Path $regKey).PSObject.Properties) {
        if ($prop.Name -notlike 'PS*') { $regNames += $prop.Name }
      }
    }
  } catch { }
}
$needed = New-Object System.Collections.Generic.HashSet[string]
foreach ($b in $blocks) {
  [void]$needed.Add($RoleDefaults[[string]$b.role].Font)
  if ([string]$b.role -eq 'signer') { [void]$needed.Add($FONT_H2) }
}
[void]$needed.Add($FONT_WEST)
[void]$needed.Add($FONT_PAGE)
$missing = @($needed | Where-Object { $fn = $_; ($installed -notcontains $fn) -and (@($regNames | Where-Object { $_ -like ($fn + '*') }).Count -eq 0) })
if ($missing.Count -gt 0) {
  $fontMsg = "未能确认以下字体已安装（文档仍按标准字体名写入，不做替代；渲染效果取决于打开文档的机器是否装有该字体）: " + ($missing -join ', ')
  if ($StrictFont) { Write-Error $fontMsg; exit 3 }
  Write-Warning $fontMsg
}

$page = $spec.page
function Get-PageVal($obj, $name, $dflt) {
  if ($null -ne $obj -and $null -ne $obj.$name) { return [double]$obj.$name }
  return [double]$dflt
}
$marginTop = Get-PageVal $page 'marginTopMm' 37
$marginBottom = Get-PageVal $page 'marginBottomMm' 35
$marginLeft = Get-PageVal $page 'marginLeftMm' 28
$marginRight = Get-PageVal $page 'marginRightMm' 26
$gridLines = Get-PageVal $page 'gridLines' 22
$gridChars = Get-PageVal $page 'gridChars' 28

$word = $null; $doc = $null
$applied = New-Object System.Collections.ArrayList
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Add()

  $normal = $doc.Styles.Item(-1)
  $normal.Font.NameFarEast = $FONT_BODY
  $normal.Font.NameAscii = $FONT_WEST
  $normal.Font.NameOther = $FONT_WEST
  $normal.Font.Size = 16

  $ps = $doc.PageSetup
  $ps.PaperSize = 7
  $ps.TopMargin = $word.CentimetersToPoints($marginTop / 10)
  $ps.BottomMargin = $word.CentimetersToPoints($marginBottom / 10)
  $ps.LeftMargin = $word.CentimetersToPoints($marginLeft / 10)
  $ps.RightMargin = $word.CentimetersToPoints($marginRight / 10)
  $ps.HeaderDistance = $word.CentimetersToPoints(2.0)
  $ps.FooterDistance = $word.CentimetersToPoints(2.8)
  $ps.LayoutMode = 1
  $ps.LinesPage = [int]$gridLines
  $ps.CharsLine = [int]$gridChars

  $doc.Content.Text = ($texts -join $CR)

  $paras = $doc.Paragraphs
  if ($paras.Count -lt $blocks.Count) { Write-Error ("段落数不符，期望 " + $blocks.Count + " 实际 " + $paras.Count); exit 5 }

  for ($i = 0; $i -lt $blocks.Count; $i++) {
    $b = $blocks[$i]
    $role = [string]$b.role
    $d = $RoleDefaults[$role]
    $p = $paras.Item($i + 1)
    $rng = $p.Range

    $fontFarEast = [string]$b.fontFarEast; if (-not $fontFarEast) { $fontFarEast = $d.Font }
    $sizePt = $d.Size; if ($null -ne $b.sizePt) { $sizePt = [double]$b.sizePt }
    $align = [string]$b.align; if (-not $align) { $align = $d.Align }
    $first = $d.First; if ($null -ne $b.indentFirstChars) { $first = [double]$b.indentFirstChars }
    $left = $d.Left; if ($null -ne $b.indentLeftChars) { $left = [double]$b.indentLeftChars }
    $right = $d.Right; if ($null -ne $b.indentRightChars) { $right = [double]$b.indentRightChars }
    $line = $d.Line; if ($null -ne $b.lineSpacingPt) { $line = [double]$b.lineSpacingPt }
    $before = $d.Before; if ($null -ne $b.spaceBeforeLines) { $before = [double]$b.spaceBeforeLines }
    $after = $d.After; if ($null -ne $b.spaceAfterLines) { $after = [double]$b.spaceAfterLines }
    $colorName = $d.Color; if ($b.color) { $colorName = [string]$b.color }
    $bold = $d.Bold; if ($null -ne $b.bold) { $bold = [bool]$b.bold }

    $f = $rng.Font
    $f.NameFarEast = $fontFarEast
    $f.NameAscii = $FONT_WEST
    $f.NameOther = $FONT_WEST
    $f.Size = $sizePt
    $f.Bold = $bold
    $f.Color = $ColorMap[$colorName]

    $pf = $p.Format
    $pf.Alignment = $AlignMap[$align]
    $pf.CharacterUnitFirstLineIndent = [single]$first
    $pf.CharacterUnitLeftIndent = [single]$left
    $pf.CharacterUnitRightIndent = [single]$right
    $pf.LineSpacingRule = 4
    $pf.LineSpacing = [single]$line
    $pf.LineUnitBefore = [single]$before
    $pf.LineUnitAfter = [single]$after
    $pf.DisableLineHeightGrid = $false

    if ($null -ne $b.pageBreakBefore -and $b.pageBreakBefore) { $pf.PageBreakBefore = $true }

    if ($role -eq 'signer') {
      $raw = [string]$b.text
      $ci = $raw.IndexOf([string][char]0xFF1A)
      if ($ci -ge 0) {
        $r2 = $p.Range
        $r2.SetRange($r2.Start + $ci + 1, $r2.Start + $raw.Length)
        $r2.Font.NameFarEast = $FONT_H2
      }
    }

    [void]$applied.Add([PSCustomObject]@{
      index = $i; role = $role; fontFarEast = $fontFarEast; sizePt = $sizePt
      align = $align; indentFirst = $first; indentLeft = $left; indentRight = $right
      lineSpacingPt = $line; spaceBeforeLines = $before
    })
  }

  # 红色分隔线：优先用显式 redLine 块，否则给标题前最后一个版头块加红色下框线
  $hasRedLine = $false
  for ($i = 0; $i -lt $blocks.Count; $i++) { if ([string]$blocks[$i].role -eq 'redLine') { $hasRedLine = $true } }
  $titleIdx = -1
  for ($i = 0; $i -lt $blocks.Count; $i++) { if ([string]$blocks[$i].role -eq 'title') { $titleIdx = $i; break } }
  if (-not $hasRedLine -and $titleIdx -gt 0) {
    $prev = $paras.Item($titleIdx)
    $prevRole = [string]$blocks[$titleIdx - 1].role
    if ($prevRole -in @('redHeader', 'docNumber', 'signer')) {
      $bd = $prev.Borders.Item(-3)
      $bd.LineStyle = 1; $bd.LineWidth = 8; $bd.Color = 255
    }
  }
  else {
    for ($i = 0; $i -lt $blocks.Count; $i++) {
      if ([string]$blocks[$i].role -eq 'redLine') {
        $bd = $paras.Item($i + 1).Borders.Item(-3)
        $bd.LineStyle = 1; $bd.LineWidth = 8; $bd.Color = 255
      }
    }
  }

  # 版记分隔线：首条粗线、末条粗线、中间细线
  $bjIdx = @()
  for ($i = 0; $i -lt $blocks.Count; $i++) { if ([string]$blocks[$i].role -in @('cc', 'printer')) { $bjIdx += $i } }
  for ($k = 0; $k -lt $bjIdx.Count; $k++) {
    $p = $paras.Item($bjIdx[$k] + 1)
    if ($k -eq 0) {
      $bd = $p.Borders.Item(-1)
      $bd.LineStyle = 1; $bd.LineWidth = 8; $bd.Color = 0
    }
    else {
      $bd = $p.Borders.Item(-3)
      $bd.LineStyle = 1; $bd.LineWidth = 6; $bd.Color = 0
    }
    if ($k -eq $bjIdx.Count - 1) {
      $bd = $p.Borders.Item(-3)
      $bd.LineStyle = 1; $bd.LineWidth = 8; $bd.Color = 0
    }
  }

  # 页码：4 号半角宋体，数字左右各一条一字线，单页右空一字、双页左空一字
  $sec = $doc.Sections.Item(1)
  $sec.PageSetup.OddAndEvenPagesHeaderFooter = $true
  $footers = @{ 1 = @{ Align = 2; Ind = 'Right' }; 3 = @{ Align = 0; Ind = 'Left' } }
  foreach ($key in @(1, 3)) {
    $ftr = $sec.Footers.Item([int]$key)
    $ftr.LinkToPrevious = $false
    $fr = $ftr.Range
    $fr.Text = [string][char]0x2014 + '  ' + [string][char]0x2014
    $fr2 = $ftr.Range
    $pos = $fr2.Text.IndexOf('  ') + 1
    $target = $ftr.Range
    $target.SetRange($fr2.Start + $pos, $fr2.Start + $pos)
    $ftr.Range.Fields.Add($target, 33) | Out-Null
    $fp = $ftr.Range.Paragraphs.Item(1)
    $fp.Format.Alignment = $footers[[int]$key].Align
    if ($footers[[int]$key].Ind -eq 'Right') { $fp.Format.CharacterUnitRightIndent = 1 }
    else { $fp.Format.CharacterUnitLeftIndent = 1 }
    $ff = $fp.Range.Font
    $ff.NameFarEast = $FONT_PAGE
    $ff.NameAscii = $FONT_PAGE
    $ff.NameOther = $FONT_PAGE
    $ff.Size = 14
    $ff.Bold = $false
  }

  $outFull = [System.IO.Path]::GetFullPath($OutputPath)
  if (-not $outFull.ToLowerInvariant().EndsWith('.docx')) { $outFull = $outFull + '.docx' }
  $fmt = 16
  $doc.SaveAs([ref]$outFull, [ref]$fmt)
  $pages = $doc.ComputeStatistics(2)
  $doc.Close(0); $doc = $null

  $report = [PSCustomObject]@{
    engine = 'A(word-com)'
    output = $outFull
    pages = $pages
    page = [PSCustomObject]@{
      size = 'A4'
      marginTopMm = $marginTop; marginBottomMm = $marginBottom
      marginLeftMm = $marginLeft; marginRightMm = $marginRight
      gridLines = $gridLines; gridChars = $gridChars
    }
    fonts = [PSCustomObject]@{
      title = $FONT_TITLE; h1 = $FONT_H1; h2 = $FONT_H2; body = $FONT_BODY
      western = $FONT_WEST; pageNumber = $FONT_PAGE
    }
    blocks = @($applied)
  }
  $json = $report | ConvertTo-Json -Depth 6
  if ($ReportPath) {
    $utf8bom = New-Object System.Text.UTF8Encoding($true)
    $utf8noBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($ReportPath), $json, $utf8noBom)
  }
  else { Write-Output $json }
  Write-Output ("已生成: " + $outFull)
  exit 0
}
catch {
  Write-Error ("排版失败（第 " + $_.InvocationInfo.ScriptLineNumber + " 行）: " + $_.Exception.Message + " @@ " + $_.ScriptStackTrace)
  exit 5
}
finally {
  if ($doc) { try { $doc.Close(0) } catch {} }
  if ($word) { try { $word.Quit() } catch {} }
}
