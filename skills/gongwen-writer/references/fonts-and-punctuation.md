# 字体矩阵、字号与中文标点

## 1. 四款方正字体的真实族名（Windows 注册名）

| 用途 | 口述名 | **Windows 真实族名（必须用这个）** |
|---|---|---|
| 标题、发文机关标志 | 方正小标宋GBK | `方正小标宋_GBK` |
| 一级标题 | 方正黑体GBK | `方正黑体_GBK` |
| 二级标题 | 方正楷体GBK | `方正楷体_GBK` |
| 正文 | 方正仿宋GBK | `方正仿宋_GBK` |

**必须写带下划线的真名**。写成"方正仿宋GBK""仿宋_GB2312""方正小标宋简体"会静默回退到默认字体，不报错，用户看不出来。同族近似字体（方正小标宋简体、方正仿宋_GB2312 等）不能替代。

其余字体：西文与阿拉伯数字 `Times New Roman`；页码例外用 `宋体`（国标 7.5 要求 4 号半角宋体）。

## 2. 中西文必须分槽设置

Word 的 `Font.Name` 一设即覆盖全部槽位（中文、西文、复杂文种），会把 Times New Roman 一起顶掉。正确做法是分别设置：

```powershell
$font.NameFarEast = "方正仿宋_GBK"   # 中文槽
$font.NameAscii  = "Times New Roman" # 西文槽
$font.NameOther  = "Times New Roman" # 其他文种槽
$font.Name       = "Times New Roman" # 不设！会覆盖中文槽
```

DOCX 层面等价于 `w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="方正仿宋_GBK"`。

页码是唯一例外：中文、西文、数字全部 `宋体`。

## 3. 字号对照（Word 磅值）

| 名称 | 磅值 | 用处 |
|---|---|---|
| 2 号 | 22pt | 标题、发文机关标志 |
| 3 号 | 16pt | 正文、一/二/三/四层序数、附件说明、成文日期、签发人 |
| 4 号 | 14pt | 版记（抄送、印发机关与日期）、页码 |
| 小 4 号 | 12pt | 版记中的其他要素（如"抄送"单独成行时） |

## 4. 角色 → 字体/字号矩阵（脚本默认值）

| role | 中文字体 | 西文/数字 | 字号 | 对齐 | 首行缩进 | 左/右缩进 | 行距 |
|---|---|---|---|---|---|---|---|
| redHeader | 方正小标宋_GBK | Times New Roman | 22pt | 居中 | 0 | 0 | 固定 29 磅 |
| docNumber | 方正仿宋_GBK | Times New Roman | 16pt | 居中（上行文左空 1 字） | 0 | 0/1 | 固定 29 磅 |
| signer | 方正仿宋_GBK（姓名楷体_GBK） | Times New Roman | 16pt | 右 | 0 | 右空 1 字 | 固定 29 磅 |
| title | 方正小标宋_GBK | Times New Roman | 22pt | 居中 | 0 | 0 | 固定 29 磅，段前空 2 行 |
| mainRecipient | 方正仿宋_GBK | Times New Roman | 16pt | 左 | 0 | 0 | 固定 29 磅 |
| body | 方正仿宋_GBK | Times New Roman | 16pt | 两端对齐 | 2 字符 | 0 | 固定 29 磅 |
| h1 | 方正黑体_GBK | Times New Roman | 16pt | 左 | 2 字符 | 0 | 固定 29 磅 |
| h2 | 方正楷体_GBK | Times New Roman | 16pt | 左 | 2 字符 | 0 | 固定 29 磅 |
| h3 / h4 | 方正仿宋_GBK | Times New Roman | 16pt | 左 | 2 字符 | 0 | 固定 29 磅 |
| attachmentNote | 方正仿宋_GBK | Times New Roman | 16pt | 左 | 0 | 左空 2 字 | 固定 29 磅，段前空 1 行 |
| issuer | 方正仿宋_GBK | Times New Roman | 16pt | 右 | 0 | 0 | 固定 29 磅 |
| date | 方正仿宋_GBK | Times New Roman | 16pt | 右 | 0 | 右空 4 字 | 固定 29 磅 |
| cc | 方正仿宋_GBK | Times New Roman | 14pt | 左 | 0 | 左右各空 1 字 | 固定值，段前空 1 行 |
| printer | 方正仿宋_GBK | Times New Roman | 14pt | 分布 | 0 | 左空 1 字 / 右空 1 字 | 固定值 |
| pageNumber | 宋体 | 宋体 | 14pt | 单页右 / 双页左 | 0 | 右空 1 字 / 左空 1 字 | 单倍 |

## 5. 标点：中文输入法标点

正文使用中文全角标点：`，。；：？！“”（）《》、……—`；年份六角括号 `〔〕`；书名号 `《》`。

### 5.1 半角标点检测规则

在**中文语境**中出现下列半角字符即为不合规：

`,` `.` `;` `:` `?` `!` `(` `)` `<` `>` `"` `'` `[` `]`

判定"中文语境"：该字符**紧邻的任一侧**是 CJK 汉字（一-鿿）或中文标点。

### 5.2 不得改动的例外

- 纯西文/数字串内部的符号：`3.14`、`1,000`、`A-1`、`10:30`、`192.168.1.1`
- 网址、邮箱、文件路径、英文缩写内部的符号
- 已经是全角的中文标点
- 西文引号 '"' 仅在中文句子内部使用全角 " "；纯西文段落保留半角

### 5.3 直引号转全角

中文句中的 `"` `'` 应成对转为 `“”` `‘’`（按出现顺序交替）。

## 6. 字体缺失时的处理

换机器可能没装方正字体。规则：**不静默替换**为近似字体（如仿宋_GB2312），先报告缺哪一款，保留文字不变，交用户决定是安装字体还是改用其他字体。

检测方式（PowerShell）：

```powershell
Add-Type -AssemblyName System.Drawing
$installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }
$missing = @('方正小标宋_GBK','方正黑体_GBK','方正楷体_GBK','方正仿宋_GBK','Times New Roman','宋体') |
    Where-Object { $installed -notcontains $_ }
```
