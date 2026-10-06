# Known gaps of the pinned build (docx)

Each gap is a strict expected failure in `tests/` (key in brackets): the day a new pin closes it, the suite
says so, and this page is updated. "Fallback" means: do that step only with the default `docx` skill
(python-docx, lxml or LibreOffice), and keep rdocx for the rest and for the verification.

A step that rdocx blocks, or gets wrong, and that this page does not list is a new gap: report it as
`SKILL.md` describes (Reporting a bug or a missing feature).

## Comments

| Gap | What happens | Workaround / fallback |
|---|---|---|
| A comment of several paragraphs [comment-several-paragraphs] | Word and Google Docs key each `w15:commentEx` row (thread and resolved state) by the `w14:paraId` of the comment's last paragraph; rdocx uses the first. A reply of several paragraphs, or a reply to a parent of several paragraphs, reads `parent_id` None; a resolved comment of several paragraphs reads `resolved` False; `reply_to` and `resolve_comment` on a parent of several paragraphs write its first paragraph's paraId, which Word does not follow. A comment text with `\n` is written as one paragraph with the newline raw in `w:t`, while `text` joins a comment's paragraphs with `\n`. Comments of one paragraph are right | read threads from the XML: map each `w15:paraId` of `word/commentsExtended.xml` to the comment of `word/comments.xml` whose last paragraph carries it (`paraIdParent` the same way). To carry such a thread over, rebuild it with `add_comment_on_text` and `reply_to`, each comment as one line of text (its paragraphs joined with a space), then check the result in the XML |

