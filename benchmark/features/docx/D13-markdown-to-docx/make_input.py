"""D13 input: a Markdown file with three heading levels, two kinds of lists, a table, a code block, emphasis and
a link."""
import sys
from pathlib import Path

FENCE = "`" * 3
NOTES = f"""# Project notes

Intro paragraph with **bold text** and *italic text* and a [project link](https://example.com/project).

## Tasks

- Collect the samples
- Label the boxes
- Ship to the lab

### Steps

1. Open the kit
2. Fill the form
3. Seal the box

## Results

| Site | Samples | Status |
|------|---------|--------|
| North | 12 | done |
| South | 9 | pending |

## Code

{FENCE}python
def total(xs):
    return sum(xs)
{FENCE}

Final paragraph.
"""
(Path(sys.argv[1]) / "notes.md").write_text(NOTES)
