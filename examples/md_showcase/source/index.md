# md_showcase — every docxsphinx feature, in one Markdown file

This example exercises every feature landed in the docxsphinx 2.0.0 work
from a single Markdown source via MyST. Rebuild with
`make example DIR=md_showcase` from the repo root, then open the
resulting `md_showcase_project-0.1.docx` in Microsoft Word or
LibreOffice to verify visually.

## Headings

### Level three

Content under level three.

#### Level four

Content under level four.

## Inline formatting

A paragraph with **bold**, *italic*, ***bold italic***, `inline code`,
and ~~strikethrough~~. Water is H<sub>2</sub>O and Einstein said
E = mc<sup>2</sup>, give or take. These inline spans should all render
with matching run-level formatting in the output.

A paragraph with mixed formatting: the **bold run**, the *italic run*,
the `monospace run`, and the ~~struck run~~ appear as separate runs
within one paragraph.

## Bullet list

- alpha
- beta
- gamma

## Numbered list

1. first step
2. second step
3. third step — with **bold** text embedded

## Nested lists (mixed types)

1. outer numbered one
   - inner bullet a
   - inner bullet b
     - deeper bullet
2. outer numbered two
   1. inner numbered 2.1
   2. inner numbered 2.2

## Task list (GFM checkboxes)

- [x] scaffold the 2.0.0 release branch
- [x] implement the ten feature gaps
- [ ] tag and push to PyPI
- [ ] announce on the project README
- plain bullet (not a task) should not carry ☐ / ☒

## Definition list

docxsphinx
: A Sphinx builder that emits Microsoft Word (.docx) from
  reStructuredText and Markdown input.

Pygments
: Syntax-highlighter used for code-block coloring. The docxsphinx
  writer consumes its tokens and emits one colored run per token.

MyST
: The Markdown parser that reaches docutils' node model so this single
  builder handles both `.rst` and `.md` sources.

## GFM pipe table

| Component | Role                              | Source              |
| --------- | --------------------------------- | ------------------- |
| Builder   | orchestration + toctree inlining  | `builder.py`        |
| Writer    | docutils visitor → python-docx    | `writer.py`         |
| Compat    | Sphinx/docutils version shims     | `_compat.py`        |
| Helpers   | raw OOXML (hyperlinks, footnotes) | `_docx_helpers.py`  |

## Fenced code blocks

### Python

```python
from __future__ import annotations


def fib(n: int) -> int:
    """Return the n-th Fibonacci number (0-indexed)."""
    a, b = 0, 1
    for _ in range(n):
        a, b = b, a + b
    return a


if __name__ == '__main__':
    for i in range(10):
        print(i, fib(i))
```

### Bash

```bash
#!/usr/bin/env bash
set -euo pipefail

for f in *.md; do
    echo "processing $f"
    wc -l "$f"
done
```

### JSON

```json
{
  "project": "docxsphinx",
  "version": "2.0.0",
  "features": ["markdown", "rst", "tasklist", "footnotes"]
}
```

### Plain (no language)

```
+---+---+---+
| o | o | o |
+---+---+---+
| x |   | x |
+---+---+---+
| o | o | o |
+---+---+---+
```

## Images

Native size:

![sample image](image1.png)

Sized and centered:

```{image} image1.png
:alt: centered 2-inch image
:width: 2in
:align: center
```

Height in cm:

```{image} image1.png
:alt: 3cm tall image
:height: 3cm
```

Inline image inside a paragraph — text then ![inline](image1.png) then
text, all flowing in the same paragraph.

## External hyperlinks

Visit the [Sphinx documentation](https://www.sphinx-doc.org/) or the
[python-docx guide](https://python-docx.readthedocs.io/). Link text may
itself contain **[bold link text](https://example.com/)**,
*[italic link text](https://example.org/)*, or
[~~struck link text~~](https://example.net/) — the inline formatting
round-trips inside the `w:hyperlink` element.

## Internal anchor links

See [the admonitions section below](#admonitions-every-type) for a
demonstration of every built-in callout.

See [the footnotes section below](#footnotes-with-inline-formatting) to
check footnote round-tripping.

(custom-target)=
## Custom anchor target

This heading has both its auto-generated id and a custom anchor defined
by the preceding `(custom-target)=` line.  A reference elsewhere can
point here via `[text](#custom-target)`.

## Admonitions (every type)

:::{note}
A short note callout with **bold** and *italic* spans and some
`inline code`.
:::

:::{warning}
Proceed with caution: the operation you are about to perform cannot
easily be undone.
:::

:::{tip}
You can use `make example DIR=md_showcase` to rebuild this sample.
:::

:::{caution}
Moving parts — keep fingers clear.
:::

:::{important}
Read the release notes before upgrading.
:::

:::{hint}
Try `make test-fast` for a faster inner loop.
:::

:::{error}
Failed to open the document.
:::

:::{danger}
Do not taunt happy-fun-ball.
:::

:::{attention}
The deprecation takes effect in v3.0.
:::

:::{admonition} Custom heading
Generic admonitions take the user's heading as the bold title run.

- Nested content preserves its structure inside the callout.
- Including bullets, code, etc.
:::

:::{note}
A multi-paragraph note.

The second paragraph demonstrates that admonition bodies can carry any
block-level content, including:

- Bullet lists inside admonitions
- Inline formatting like **bold** and `code`

```python
# Even a code block inside an admonition.
print("hello from inside a note")
```
:::

## Figures with captions

```{figure} image1.png

The pair of cats, figure one. Captions auto-number via Word `SEQ`
fields.
```

```{figure} image1.png

Dogs playing, figure two. The number increments automatically.
```

## Footnotes with inline formatting

A sentence with a plain footnote reference[^plain] and another with
embedded **bold**[^richtext].

[^plain]: This is a plain footnote body.

[^richtext]: This footnote body contains **bold**, *italic*, and
  ~~struck~~ inline formatting — all three preserve their run-level
  formatting inside the `word/footnotes.xml` part.

## Blockquote

> Blockquotes flow through as a block of indented content. The
> docxsphinx writer currently renders them as normal paragraphs;
> full blockquote styling is a future enhancement.

## Horizontal rule

---

That's the end of the showcase. Open the generated `.docx` in Word or
LibreOffice and confirm that everything above renders as expected.
