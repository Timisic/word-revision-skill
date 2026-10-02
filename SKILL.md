---
name: word-revision
description: Edit existing Word .docx files with tracked text changes and anchored comments, preserve earlier review history, and verify the resulting edits. Use for local text replacements, review comments, and revision checks.
---

# Word Revision

Turn editing requests into verifiable Word revisions. Use the bundled script to read documents, apply local edits, and check text, review history, and formatting structure.

## Workflow

1. **Read the current file.** Run `scripts/word_revision.py inspect` from the skill directory. Use `--query` or `--paragraph` to retrieve relevant text, and read the comments and revisions associated with the task. See [references/commands.md](references/commands.md) for command options and the plan format.
2. **Prepare a local edit plan.** Use the returned file fingerprint, paragraph IDs, and unique source passages. Use the requested review author, or `Reviewer` when unspecified. Preserve paragraph and table structure, expressing changes as minimal text differences.
3. **Preflight the entire batch.** `check` simulates the edits and verifies the result in memory. If it reports a targeting conflict, refine the source anchor or plan using the reported location, then check the entire batch again.
4. **Create a revised copy.** `apply` writes a new `.docx` and verification receipt. Deliver to the requested location. Before replacing the original, confirm that the editor has saved and closed the target and that its disk fingerprint still matches the plan, then replace it with the verified copy.
5. **Read back and verify.** `verify` checks the output against the original, edit plan, and receipt. Deliver the revised document with a brief account of the edits, author, preserved review history, and verification results.

Complete these steps continuously for text edits the user has already authorized. The plan and receipt support execution and traceability; keep the delivery message concise.

## Choose the verification method

For ordinary text edits and comments, use programmatic checks: planned text matches, existing revisions and comments are preserved, unrelated parts remain intact, and undoing only this batch restores the previous text and formatting structure.

When the task includes layout requirements such as image placement, column widths, pagination, or font substitution, or a structural check identifies a concrete layout concern, render and inspect the relevant pages. Use the target Word environment's layout for exact page or line numbers.

Report content and structure, layout, and native Word verification separately. See [references/verification.md](references/verification.md) for what each check establishes.

## Editing conventions

- Preserve existing revisions and comments as review history. New text inherits the first affected run's character formatting; retained text keeps its original formatting.
- Evaluate suggestions in comments against the user's task before adding them to the edit plan.
- Preflight identifies targets with structural conflicts, such as fields or existing revision boundaries. Choose a Word operation appropriate to that structure and the user's intent, then perform equivalent verification before delivery.
- Deliver text-focused tasks once programmatic verification passes, reporting layout as unchecked when applicable.

Requirements: Python 3.10+ and `lxml`. See [README.md](README.md) for installation and examples.
