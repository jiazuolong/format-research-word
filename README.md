# 科研报告 Word 排版工作流

整理中文材料科研报告的 Word 排版。项目包括一套可供 Codex 执行的完整工作流、可编辑的 Word 模板，以及在本地运行的 Python 命令行工具。支持 Windows、macOS 和 Linux 上的 Python 3.10 或更高版本。

**选择入口**

| 你要做什么 | 入口 |
| --- | --- |
| 查看排版规则，或交给同事执行 | [Word 版标准工作流](docs/workflow-zh.docx) |
| 统一正文、三级标题、题注和表格的样式 | [可编辑 Word 样式模板](assets/report-styles.docx) |
| 上传报告，请 Codex 按模板排版或续接章节 | [Codex 技能](SKILL.md) |
| 批量处理可编辑文本，盘点静态引文和长小数 | 下方的本地命令行 |

## 本地运行

在终端中克隆项目并安装。建议使用 Python 虚拟环境。

```sh
git clone https://github.com/jiazuolong/format-research-word.git
cd format-research-word
python -m pip install .
format-research-word audit "报告.docx" --json "排版检查.json"
format-research-word format "报告.docx" "报告_排版版.docx"
```

希望同步重排静态方括号引文时，先检查 JSON 中的编号映射、未引用条目和待判定标记，再将下列命令作为排版步骤。输出路径应使用新文件名。

```sh
format-research-word format "报告.docx" "报告_引文排版版.docx" --renumber-citations
```

`python -m format_research_word` 与已安装的命令等效。仅需盘点文献与数字时，也可直接从仓库运行 `python scripts/audit_report.py "报告.docx"`，这条命令只使用 Python 标准库。文档处理在本机完成。

## 排版规则

| 项目 | 示例与处理 |
| --- | --- |
| 化学式 | 识别常见元素组合和组分表达式。元素与计量数使用 Times New Roman，计量数和化学式中的组分系数使用 Word 原生下标。 |
| 变量 | 组分变量 x 用正体，保留乘号 ×。 |
| 数值 | 小数部分达到五位时保留三位有效数字。−4.497105916 → −4.50，0.01870 → 0.0187。四位小数 10.9415 保留原值。 |
| 空格 | “在 300K 下，用 DFT 计算”变为“在300 K下，用DFT计算”。数值和物理单位间保留一个空格，百分号同样处理。 |
| 参考文献 | 若第一处是 [3,1]，两篇文献依次获得新号 1、2。后续引用复用新号；文末整条移动，未引用文献按原相对顺序追加。 |
| 模板与编号 | 完整工作流依目标模板处理正文、大纲、图表题注、公式号、附录及交叉引用，最后在 Word 中重新分页并刷新目录。 |

### 命令行的处理范围

`format` 修改主文档内的常规文本段落、图表题注和表格文字，可跨相邻文本片段识别化学式；处理原生数学公式中单个数学文本片段里的长小数、组分变量 x 和微分 dx。图件和嵌入对象随 DOCX 保存。它沿用输入文档的页面和段落样式。章节合并、指定模板套用、复杂公式、图片内文字、脚注和最终目录分页，按 [完整工作流](SKILL.md) 在 Word 或文档代理中处理。

`--renumber-citations` 适用于正文使用方括号、文后每条记录独占一段且有可编辑编号的单一书目。组合引用和区间会展开，按新号排序。整数坐标与引文可能写成相同的方括号形式，需要结合上下文判断。含 Zotero、EndNote 等动态引用、脚注引用、修订记录或歧义编号时，请使用原文献管理机制与工作流。检查报告中的段落号按文档读取顺序计算，包含表格中的段落。

## 在 Codex 中使用

将整个仓库作为名为 `format-research-word` 的技能目录放入当前环境的技能目录，例如 `$CODEX_HOME/skills/format-research-word`。保留其中的 `assets`、`scripts` 和 `src`。在能识别该技能的会话中上传报告，输入：

> 使用 $format-research-word 按标准工作流整理附件，核对化学式、数值、单位、引文和目录，另存为新版。

续接两份报告时，说明底稿、来源、起始标题和插入位置。另附格式模板时，以指定模板为准。

## 开发与验证

```sh
python -m pip install -e .
python -m unittest discover -s tests -v
```

测试使用合成报告，验证数值精度、原生上下标、跨文本片段的化学式、文献映射和 DOCX 对象保留。项目采用 [MIT 许可证](LICENSE)。
