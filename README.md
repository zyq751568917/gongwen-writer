# gongwen-writer — 党政机关公文写作与排版 Skill

一个可安装到任意 Mavis / MiniMax Code Agent 的公文技能包：按 **GB/T 9704-2012《党政机关公文格式》** 撰写、排版、校验党政机关公文，默认交付 Word。

## 能做什么

- **读**：提取 Word / Excel / PPT / PDF 材料文字（PDF 走 Word 重排，无文本层时明确提示）
- **写**：按 15 种法定文种的结构模板起草正文，层次序数与字体对应国标，语言书面庄重
- **排**：Word 成品全套版式——版心 37/35/28/26mm、每页 22 行每行 28 字、红色分隔线、发文字号、签发人、标题、主送机关、附件说明、成文日期、版记、页码"一 N 一"
- **校**：逐段输出字体/字号/缩进/行距/页面/页码/标点合规报告，FAIL 清零才交付
- **依据**：结构存疑时检索官媒与政府官网范例，并给出可点链接；查不到就如实说明，不编造条款号

## 字体要求（Windows）

| 用途 | 字体 |
|---|---|
| 标题、发文机关标志 | `方正小标宋_GBK` |
| 一级标题 | `方正黑体_GBK` |
| 二级标题 | `方正楷体_GBK` |
| 正文 | `方正仿宋_GBK` |
| 西文与阿拉伯数字 | `Times New Roman` |
| 页码 | `宋体`（国标 7.5 例外） |

四款方正 GBK 字体需自行安装。缺失时脚本会报告并保留原文，不会静默替换。

## 安装

把 `skills/gongwen-writer/` 整个目录复制到目标 Agent 的技能目录即可：

| Agent | 技能目录 |
|---|---|
| MiniMax Code（本机） | `C:/Users/Administrator/.minimax/skills/gongwen-writer/` |
| 其他兼容目录 | `<agent 配置目录>/skills/gongwen-writer/` |

安装后新建会话即可用"公文""请示""按公文格式改"等说法触发。

## 依赖

| 能力 | 用途 | 缺失时 |
|---|---|---|
| Microsoft Word（COM） | 主排版引擎、校验、导出 PDF | 自动降级 Node 兜底引擎，红线/版记/页码一字线需在 Word 中补齐 |
| Node.js 18+ | 兜底引擎、标点校验 | 兜底不可用，只交付纯文本与版式参数 |
| 四款方正 GBK 字体 | 中文版式 | 报告缺失字体，交用户决定 |

不需要 Python、LibreOffice 或任何 npm 第三方包。

## 用法

技能被加载后，Agent 会走"读取 → 定结构并给依据 → 写正文 → 生成 spec.json → 排版 → 标点规范化 → 校验 → **交付前用问答确认要素** → 交付 docx"的闭环。

命令行直接用（以 PowerShell 为例）：

```powershell
powershell -ExecutionPolicy Bypass -File skills/gongwen-writer/scripts/format-gongwen.ps1 -SpecPath spec.json -OutputPath 请示.docx
powershell -ExecutionPolicy Bypass -File skills/gongwen-writer/scripts/verify-gongwen.ps1 -Path 请示.docx -JsonPath 校验.json
node skills/gongwen-writer/scripts/ooxml-format.mjs --spec spec.json --out 请示.docx
```

## 目录结构

```
skills/gongwen-writer/
├─ SKILL.md                 # 执行规则
├─ references/              # 国标条款、字体矩阵、15 文种、语言规范、依据策略、spec 契约
├─ scripts/                 # 抽取/排版/校验/标点/PDF/兜底引擎
└─ examples/spec-sample.json
```

## 边界

不生成红头图片与电子签章，不做扫描件 OCR，不替代本单位的保密审查与合法性审查，不推送任何远端仓库。

## 许可

MIT，见 [LICENSE](LICENSE)。国标条款与官方范例引用仅供格式参考，公文内容由使用者负责。
