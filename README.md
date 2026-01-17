# DocxCommentInjector（线性化批注/修订/高亮）

把 `.docx` 里的"批注/修订/高亮"等边栏信息，转换为 **Markdown 格式**的可线性阅读文本，便于 AI 直接理解"原文 + 评论/建议 + 改动"。

## 一键运行（需要安装 [uv](https://docs.astral.sh/uv/)）

无需克隆，直接从 GitHub 运行：

```bash
uvx --from git+https://github.com/fanxing-6/DocxCommentInjector docx-linearize input.docx output.md
```

输出到 stdout：

```bash
uvx --from git+https://github.com/fanxing-6/DocxCommentInjector docx-linearize input.docx -
```

## 本地开发

克隆后运行：

```bash
uv run docx-linearize input.docx output.md
```

## 输出格式

### Markdown 结构

- **标题层级**：识别 Word 标题样式，转换为 `#`、`##`、`###` 等
- **列表**：识别有序/无序列表，转换为 `1.` 或 `-`
- **粗体/斜体**：`**粗体**`、`*斜体*`
- **表格**：转换为 Markdown 表格格式

### 修订与高亮标记

- **插入**：`{+插入文本+}`
- **删除**：`[-删除文本-]`
- **高亮**：`==高亮文本==`（自动合并连续高亮，避免 `====` 碎片）

### 批注格式（块级）

批注以独立的 Markdown 引用块呈现，包含原文范围：

```markdown
正文内容...

> **[批注 #1]** 作者 (2024-01-15)
> **原文**：被批注的文本片段
> **批注**：批注的具体内容

下一段正文...
```

这种格式的优点：
- 批注与原文对应关系清晰
- 嵌套批注各自独立，不会互相干扰
- AI 能直接理解"哪段原文对应哪条批注"

## 技术原理

- 直接解析 `.docx`（本质是 zip）中的 XML 文件
- 不依赖 python-docx 等高级抽象层，直接遍历 WordprocessingML 节点
- 使用区间跟踪机制收集批注范围内的原文
- 识别 `styles.xml` 中的标题样式和 `numbering.xml` 中的列表定义
