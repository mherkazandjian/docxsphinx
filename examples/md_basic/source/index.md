# md_basic — Markdown sanity demo

This example demonstrates that `docxsphinx` can render a plain Markdown
source file to a Word document via `myst-parser`, with no additional
configuration beyond enabling MyST as the `.md` source parser.

## Paragraphs and inline formatting

Here is a paragraph with **bold text**, *italic text*, ***bold italic***,
and some inline `code literal`. Markdown also supports ~~strikethrough~~
text when the `strikethrough` MyST extension is enabled.

A second paragraph, separated by a blank line, with a trailing sentence
that wraps across multiple soft-break lines
but still becomes one paragraph in the output.

## Nested headings

### Level three

Content under a level-three heading.

#### Level four

Content under a level-four heading.

## A short list

- alpha
- beta
- gamma

## A short numbered list

1. first
2. second
3. third
