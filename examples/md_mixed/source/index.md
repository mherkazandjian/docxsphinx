# md_mixed — Markdown and reStructuredText side-by-side

This is the Markdown master document for a project that also includes
a reStructuredText file. Both are pulled into the single output `.docx`
via the `toctree` below.

## Why mix formats?

Teams migrating gradually from RST to MD often need both to coexist.
`docxsphinx` is format-agnostic — the builder operates on the docutils
doctree produced by whichever parser Sphinx picks per file extension.

```{toctree}
:maxdepth: 2
:caption: Contents

appendix_rst
```

## A Markdown bullet list

- first MD bullet
- second MD bullet
- third MD bullet
