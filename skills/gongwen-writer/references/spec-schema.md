# spec.json 契约

排版脚本的唯一输入。内容与角色的判断由模型完成，脚本只按契约执行，因此 **spec.json 必须自洽**：每个 block 一个角色，text 为该段的最终文字。

## 1. 顶层结构

```json
{
  "meta": {
    "title": "关于××××的请示",
    "issuer": "××××有限公司",
    "date": "2026年11月25日",
    "docNumber": "锡政字〔2026〕15号",
    "isUpward": true,
    "redHeader": true,
    "pageNumber": true
  },
  "page": {
    "marginTopMm": 37,
    "marginBottomMm": 35,
    "marginLeftMm": 28,
    "marginRightMm": 26,
    "gridLines": 22,
    "gridChars": 28
  },
  "blocks": [ { "role": "title", "text": "…" } ]
}
```

- `meta`：仅作记录与脚本参数（如上行文决定发文字号位置），**不驱动排版**；排版只认 `blocks`。
- `page`：可省略，省略时用国标默认值（37/35/28/26，22 行 28 字）。

## 2. block 字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `role` | string | 必填，见角色表；未知角色脚本报错退出 |
| `text` | string | 必填，该段最终文字（可含 `\n` 换行，渲染为换行符） |
| `align` | string | `left` / `center` / `right` / `justify`，覆盖角色默认 |
| `fontFarEast` | string | 覆盖中文字体族名 |
| `sizePt` | number | 覆盖字号（磅） |
| `bold` | bool | 默认 false |
| `color` | string | `red` / `black` / `auto`，默认 auto |
| `indentFirstChars` | number | 首行缩进字符数 |
| `indentLeftChars` | number | 左缩进字符数 |
| `indentRightChars` | number | 右缩进字符数 |
| `lineSpacingPt` | number | 行距磅值，默认 29 |
| `spaceBeforeLines` | number | 段前空行数 |
| `spaceAfterLines` | number | 段后空行数 |
| `pageBreakBefore` | bool | 段前分页（附件独立页用） |

## 3. 角色表（脚本默认值见 fonts-and-punctuation.md 第 4 节）

| role | 含义 |
|---|---|
| `redHeader` | 发文机关标志（红头） |
| `redLine` | 红色分隔线（脚本自动在 redHeader/docNumber 之后插入，也可显式给出） |
| `docNumber` | 发文字号 |
| `signer` | 签发人（上行文） |
| `title` | 公文标题 |
| `mainRecipient` | 主送机关 |
| `body` | 正文段落 |
| `h1` | 一级标题"一、" |
| `h2` | 二级标题"（一）" |
| `h3` | 三级标题"1." |
| `h4` | 四级标题"（1）" |
| `attachmentNote` | 附件说明 |
| `issuer` | 发文机关署名 |
| `date` | 成文日期 |
| `cc` | 抄送机关（版记） |
| `printer` | 印发机关与印发日期（版记） |
| `note` | 附注（成文日期下空一行左空 2 字加圆括号） |
| `attachmentTitle` | 附件标题（居中，版心第三行） |
| `attachmentMarker` | "附件"标识（左上角顶格，3 号黑体） |

## 4. 校验规则（脚本执行）

- 未知 `role` → 非零退出，列出未知角色名。
- `blocks` 为空 → 非零退出。
- `text` 缺失或为空字符串 → 非零退出，列出块序号。
- `page` 中的边距为负或网格行数非正 → 非零退出。
- 同一 `text` 中的 `\n` 渲染为软换行（同一段内换行），不拆段。
