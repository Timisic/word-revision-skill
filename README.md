# Word Revision

[中文](#中文) · [English](#english)

## 中文

面向 AI Agent 的 Word 修订技能。读取现有 `.docx` 的正文、批注和审阅历史，通过脚本批量生成原生修订，并核对实际修改。普通文字任务默认使用程序验收；涉及版式时按需检查页面。

### 能做什么

- 读取正文、表格文字、批注、修订及字段信息。
- 跨文字片段进行局部替换，保留未改文字的格式。
- 按指定作者生成插入、删除修订，并为修改添加批注。
- 保留旧修订和批注，整批预检通过后生成修订副本。
- 对照原稿与计划核验结果，包括仅撤销本轮修改后的恢复检查。

### 使用

需要 Python 3.10+ 和 `lxml`。在已配置的 Python 环境中运行：

```bash
git clone https://github.com/Timisic/word-revision-skill.git
cd word-revision-skill
python -m pip install -r requirements.txt
```

将仓库目录作为技能安装到你的 Agent 中。Codex 可将其复制或链接到 `~/.codex/skills/word-revision`。加载后可以直接说：

> 使用 $word-revision，把这份 Word 中“重复30次评分”改为“重复10次评分”，作者为“审阅者”，保留旧修订，并验证结果。

也可独立使用命令行：

```bash
python scripts/word_revision.py inspect input.docx --query "30次"
python scripts/word_revision.py check input.docx --plan plan.json
python scripts/word_revision.py apply input.docx --plan plan.json \
  --out revised.docx --receipt receipt.json
python scripts/word_revision.py verify input.docx revised.docx \
  --plan plan.json --receipt receipt.json
```

当前核心处理正文及表格单元格内的普通文字。结构冲突会返回具体位置。核验结果分别记录内容与结构、原生 Word 行为及版式的检查状态。

- [技能说明](SKILL.md)
- [修改计划与命令示例](references/commands.md)
- [验证方式](references/verification.md)

### 测试

```bash
python -m unittest discover -s tests -v
```

测试使用合成文档。当前版本在 macOS 上通过 25 项自动测试和独立场景测试，并在 Word 16.109.1 中实测接受／拒绝本轮修订：正文和表格文字正确，旧作者修订与批注保留。GitHub Actions 配置覆盖 macOS、Linux、Windows，运行结果见仓库 Actions。

## English

Word Revision is an agent skill for verifiable edits to existing `.docx` files. It creates native tracked text changes and anchored comments, preserves earlier review history, and checks each batch against its source and edit plan. Text edits use programmatic verification by default; layout checks run when the task calls for them.

### Features

- Read body text, table text, comments, revisions, and field information.
- Replace text across runs while preserving the formatting of retained text.
- Create tracked insertions and deletions with a specified review author.
- Add comments to a selected passage or directly to a replacement.
- Preflight the entire batch, write a revised copy, and verify the result.
- Check that undoing only the current batch restores the original text and formatting structure.

### Get started

Requires Python 3.10+ and `lxml`. Clone this repository and install `requirements.txt` in your Python environment using the commands above. Install the repository as a skill in your agent; for Codex, copy or link it to `~/.codex/skills/word-revision`.

> Use $word-revision to change “30 repeated ratings” to “10 repeated ratings”, use “Reviewer” as the author, preserve existing revisions, and verify the result.

The four CLI commands are `inspect`, `check`, `apply`, and `verify`. A plan contains the source fingerprint, review author, and a list of operations. For example:

```json
{
  "source_sha256": "fingerprint returned by inspect",
  "author": "Reviewer",
  "operations": [
    {
      "type": "replace",
      "paragraph": "p000001",
      "before": "30 repeated ratings",
      "after": "10 repeated ratings",
      "comment": "Updated to match the analysis record."
    }
  ]
}
```

The current core edits ordinary text in body paragraphs and table cells. Conflicts return the affected operation and reason before any document is written. Verification reports distinguish content and structure checks, native Word behavior, and page layout checks.

Run `python -m unittest discover -s tests -v` for the synthetic regression suite. The GitHub Actions matrix covers macOS, Linux, and Windows.

---

依据公开 WordprocessingML 资料独立实现，使用 `lxml` 处理 XML。参考资料见 [验证说明](references/verification.md)。

Implemented independently from public WordprocessingML documentation, using `lxml` for XML processing. Source references are listed in the verification guide.
