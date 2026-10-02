---
name: word-revision
description: Edit existing Word .docx files with tracked text changes and anchored comments, preserve earlier review history, and verify the resulting edits. Use for Word 修订、保留旧修订、批注、局部文字替换 and revision checks.
---

# Word Revision

把修改意图转成可核验的 Word 修订。通过随附脚本读取文档、执行局部修改，并核对正文、审阅历史和格式结构。

## 工作流程

1. **读取当前文件。** 从本技能目录运行 `scripts/word_revision.py inspect`。根据任务使用 `--query` 或 `--paragraph` 获取相关文字，读取需要处理的批注和旧修订。命令参数及计划格式见 [references/commands.md](references/commands.md)。
2. **形成局部修改计划。** 使用返回的文件指纹、段落编号和唯一原文片段填写计划。沿用用户指定的修订作者；作者未指定时使用 `Reviewer`。保留原段落和表格结构，用最小文字差异表达改动。
3. **整批预检。** `check` 在内存中模拟修改并验证结果。若出现定位冲突，依据错误中的具体位置补足原文锚点或调整计划，再检查整批。
4. **生成修订副本。** `apply` 写入一个新的 `.docx` 和核验回执。按用户要求的位置交付；覆盖原件前先确认编辑器已保存并关闭目标文件、磁盘指纹仍与计划一致，再以已验证的副本替换。
5. **独立读回。** `verify` 用原稿、修改计划和回执核对输出。交付修订文档，简短说明修改范围、作者、旧修订保留情况及验证结果。

用户已经授权的文字修改可连续完成以上流程。修改计划和核验回执用于执行及追溯，交付说明保持简洁。

## 验收选择

普通文字修改与批注使用程序验收：计划文字一致、既有修订与批注保留、未涉及部件保持、仅撤销本轮修改后恢复修改前的文字及格式结构。

当任务包含图片位置、列宽、分页、字体替换等版式要求，或结构检查发现具体版式疑点时，渲染并检查相关页面。若任务需要准确页码或行号，采用目标 Word 环境的排版结果。

结果中分别说明内容与结构、版式、原生 Word 行为的验证状态。程序检查的具体含义见 [references/verification.md](references/verification.md)。

## 执行约定

- 旧修订和批注作为输入历史保留；新文字继承首个改动片段的字符格式，未改文字保持原格式。
- 批注中的建议先结合用户任务判断，再纳入修改计划。
- 字段、旧修订边界等结构冲突由预检返回具体目标；结合用户意图选择适合该结构的 Word 操作方式，完成同等核验后交付。
- 内容型任务完成程序验收即可交付；按需附上“版式未检查”的状态。

运行要求：Python 3.10+ 和 `lxml`。安装方法、示例及开源说明见 [README.md](README.md)。
