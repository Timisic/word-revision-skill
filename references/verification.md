# 验证层次

## 内容与结构

每批修改自动检查：

- 输出的当前文字逐段等于原文加修改计划。
- 仅撤销本批插入、删除和新增批注后，文档恢复原稿的规范化文字及格式结构。
- 既有修订元素及批注内容保留，新增批注的起点、终点、引用各有一个。
- 新修订标识唯一，署名匹配计划，修订开关启用。
- 修改范围限于正文与所需的修订设置、批注、关系、内容类型部件；其他部件字节一致。
- 从磁盘重新读入结果，再按计划和记录时间重建，逐部件核对，覆盖文字、批注、作者、关系及设置。

规范化只合并字符格式及 run 属性相同的相邻普通文字片段，使 Word 允许的片段拆分差异不影响比较。原有复杂元素、格式属性和批注位置仍参加结构比较。

回执记录输入/输出指纹、操作范围、修订标识、检查结果及日期。保留原稿、计划和回执，可以重现本批核验。

## 原生行为与版式

程序验收验证受支持操作的内容及结构；`full_schema`、`native_word` 与 `layout` 分别记录是否执行完整 OOXML schema、Word 原生操作和页面检查。

涉及新的编辑结构时，可用合成样例在 Word 中打开，实际执行接受/拒绝并保存读回，核对当前文字、历史修订及格式。这是扩充引擎支持范围时的回归测试。

版式任务采用目标渲染环境检查修改页及受分页影响的后续页面。对于稳定、未变的页面可复用基线结果。批注核验依靠其文字、锚点及关系，页面截图仅作为显示效果的补充。

## 回归测试

```bash
python -m unittest discover -s tests -v
```

测试使用合成 DOCX，覆盖跨文字片段、中英混排、格式继承、旧修订、表格、批注、歧义、整批失败、文件变化及输出被改动。CI 在 macOS、Linux、Windows 上运行相同测试。

## 公开依据

独立实现依据 WordprocessingML 的插入、删除和批注结构，依赖 `lxml` 的 XML 读写能力。

- [Microsoft — 文档结构及修订元素](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)
- [Microsoft — 批注及锚点](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document)
- [Microsoft — Open XML 验证](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-validate-a-word-processing-document)
- [lxml 许可证](https://lxml.de/FAQ.html#what-is-the-license)
