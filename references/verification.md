# Verification layers

## Content and structure

Each batch checks that:

- The output's current text matches the source plus the edit plan, paragraph by paragraph.
- Undoing only this batch's insertions, deletions, and new comments restores the source's normalized text and formatting structure.
- Existing revision elements and comment contents are preserved, and each new comment has one start marker, end marker, and reference.
- New revision IDs are unique, authors match the plan, and change tracking is enabled.
- Changed parts are limited to the main document and the required revision settings, comments, relationships, and content types; other parts remain byte-identical.
- Reading the output from disk and rebuilding from the plan and recorded timestamp produces matching parts, covering text, comments, authors, relationships, and settings.

Normalization merges only adjacent ordinary text runs with identical character formatting and run attributes, so equivalent run splitting does not affect comparison. Existing complex elements, formatting properties, and comment positions remain part of the structural comparison.

The receipt records input and output fingerprints, operation ranges, revision IDs, check results, and the timestamp. Retaining the original, plan, and receipt makes the batch verification reproducible.

## Native behavior and layout

Programmatic checks verify the content and structure of supported operations. The `full_schema`, `native_word`, and `layout` fields separately record whether full OOXML schema validation, native Word operations, and page inspection were performed.

When adding support for a new editing structure, open a synthetic example in Word, actually accept or reject the changes, save it, and read it back to check text, earlier revisions, and formatting. Use this as a regression test when extending the engine.

For layout tasks, use the target rendering environment to inspect changed pages and subsequent pages affected by pagination. Reuse baseline results for stable, unchanged pages. Verify comments through their text, anchors, and relationships; page images supplement checks of their appearance.

## Regression tests

```bash
python -m unittest discover -s tests -v
```

Synthetic DOCX tests cover edits across runs, mixed Chinese and English text, formatting inheritance, existing revisions, tables, comments, ambiguous anchors, whole-batch failure, source changes, and output tampering. CI runs the same tests on macOS, Linux, and Windows.

## Public references

The implementation independently follows WordprocessingML insertion, deletion, and comment structures, using `lxml` for XML processing.

- [Microsoft — Document structure and revision elements](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)
- [Microsoft — Comments and anchors](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document)
- [Microsoft — Open XML validation](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-validate-a-word-processing-document)
- [lxml license](https://lxml.de/FAQ.html#what-is-the-license)
