# Commands and edit plans

Run these examples from the skill directory. Document paths can point to any working directory.

```bash
python scripts/word_revision.py inspect input.docx --query "30 repeated ratings"
python scripts/word_revision.py inspect input.docx --paragraph p000012
python scripts/word_revision.py inspect input.docx --summary
```

`inspect` returns `source_sha256`, body paragraph IDs, current text, a text projection before insertions and deletions, comments, revision authors, and fields. Table cells also use paragraph IDs. Headers, footers, footnotes, and endnotes are listed separately in `other_stories`.

`original_text` is a projection of text insertions and deletions; paragraph merges, moves, and formatting changes remain separate revision records. Use current text and a unique anchor to locate edits.

## Plan format

```json
{
  "source_sha256": "full fingerprint returned by inspect",
  "author": "Reviewer",
  "operations": [
    {
      "type": "replace",
      "paragraph": "p000012",
      "before": "30 repeated ratings",
      "after": "10 repeated ratings",
      "comment": "Updated to match the analysis record."
    },
    {
      "type": "comment",
      "paragraph": "p000014",
      "quote": "the model performs consistently",
      "text": "Please provide a metric supporting this claim."
    }
  ]
}
```

Each operation can specify its own `author`. A replacement can include `comment` to anchor an explanation to the same edit; a standalone comment uses `type: comment`. Anchors must be unique within the specified paragraph. All operations use source-document coordinates, with mutually nonoverlapping target ranges. Multiple replacements in one paragraph can be submitted together.

`replace` generates `w:ins` and `w:del` from minimal text differences. The example above changes only “3” to “1”. Use an empty `after` string for deletion. For insertion, use existing text as the anchor and include the added text in `after`. New comments span the full `quote` range and support multiline comment text.

Replacements across ordinary text runs preserve each deleted fragment's formatting. New text inherits the first affected run's formatting for that difference. The script generates dates in UTC; the plan supplies the review author.

## Execution

```bash
python scripts/word_revision.py check input.docx --plan plan.json
python scripts/word_revision.py apply input.docx --plan plan.json \
  --out revised.docx --receipt receipt.json
python scripts/word_revision.py verify input.docx revised.docx \
  --plan plan.json --receipt receipt.json
```

All commands return JSON. Success exits with `0`; conflicts or failures exit with `2`, with `conflicts` identifying affected operations and reasons. `check` builds and verifies the entire batch in memory. `apply` publishes the completed copy and receipt to fresh paths, preserving the source file.

## Supported text structures

The core handles ordinary text in body paragraphs and table cells, including replacements, insertions, deletions, and comments across runs. Preservation checks cover existing revisions, comments, images, and other ZIP parts elsewhere in the source document.

When a source anchor includes fields, content controls, existing revisions, formatting revisions, or crosses bookmark or comment boundaries, the script returns a structural conflict and leaves the entire batch unwritten. Determine the intended review behavior at that location before choosing a native Word operation. Ordinary text anchors can be narrowed and preflighted again.

Save an open document to disk before checking it. Before writing back to the original path, close the target document and confirm its fingerprint. Retain the original, plan, and receipt alongside the revised copy to rerun `verify`.
