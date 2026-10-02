# 命令与修改计划

在技能所在目录运行以下示例。实际文档路径可以位于任何工作目录。

```bash
python scripts/word_revision.py inspect input.docx --query "30次"
python scripts/word_revision.py inspect input.docx --paragraph p000012
python scripts/word_revision.py inspect input.docx --summary
```

`inspect` 返回 `source_sha256`、正文段落编号、当前文字、插入/删除之前的文字投影、批注、修订作者和字段。表格单元格也使用段落编号。页眉页脚及脚注、尾注单独列在 `other_stories`。

`original_text` 是文字插入/删除的投影；段落合并、行移动和格式修订另外保留在修订记录中。定位主要使用当前文字与唯一锚点。

## 计划格式

```json
{
  "source_sha256": "inspect 返回的完整指纹",
  "author": "Reviewer",
  "operations": [
    {
      "type": "replace",
      "paragraph": "p000012",
      "before": "重复30次评分",
      "after": "重复10次评分",
      "comment": "重复次数已按分析记录更正。"
    },
    {
      "type": "comment",
      "paragraph": "p000014",
      "quote": "模型表现稳定",
      "text": "请补充支持稳定性判断的指标。"
    }
  ]
}
```

各操作可设置自己的 `author`。替换操作可附带 `comment`，将说明锚定在同一修改范围；独立批注使用 `type: comment`。锚点须在指定段落中唯一；同一批操作采用原稿坐标，并保持目标区间互不重叠。多个替换位于同一段落时可一次提交。

`replace` 以最小文字差异生成 `w:ins` 和 `w:del`。例如上述替换只把“3”改成“1”。删除使用空字符串 `after`；插入可用原文作为锚点，将新文字加入 `after`。新批注锚定 `quote` 的完整范围，支持多行批注正文。

跨多个普通文字片段的替换会保留被删除片段各自的格式；新增文字继承该次差异首个受影响片段的格式。日期由脚本以 UTC 生成，审阅作者由计划提供。

## 运行

```bash
python scripts/word_revision.py check input.docx --plan plan.json
python scripts/word_revision.py apply input.docx --plan plan.json \
  --out revised.docx --receipt receipt.json
python scripts/word_revision.py verify input.docx revised.docx \
  --plan plan.json --receipt receipt.json
```

所有命令输出 JSON。成功退出码为 `0`；冲突或失败为 `2`，`conflicts` 列出具体操作及原因。`check` 在内存中完成整批构建和核验。`apply` 发布完成的副本和回执，原稿保持原样；输出使用新路径。

## 当前可执行的文字结构

核心处理正文和表格单元格内的普通文字，支持跨 run 的替换、插入、删除与批注。原稿其他位置的旧修订、批注、图片及其他 ZIP 部件均纳入保留核验。

原文锚点包含字段、内容控件、旧修订、格式修订或跨越书签/批注边界时，脚本返回结构冲突并保持整批未写入。收到此类结果后，先明确该位置需要怎样的审阅语义，再选原生 Word 操作；普通文本锚点可直接缩小后重新预检。

打开中的文档先保存到磁盘，再执行检查；写回原路径之前关闭目标文档并核对文件指纹。交付副本时保留回执和原稿即可再次运行 `verify`。
