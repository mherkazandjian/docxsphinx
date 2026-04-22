# md_links — links and cross-references

## External URLs

Visit the [Sphinx documentation](https://www.sphinx-doc.org/) to learn more,
or read the [python-docx guide](https://python-docx.readthedocs.io/) for
output-side details.

The repository lives at
[github.com/mherkazandjian/docxsphinx](https://github.com/mherkazandjian/docxsphinx).

## Reference-style links

Here is a link to the [MyST project][myst], which powers the Markdown
input path used by this sample.

[myst]: https://myst-parser.readthedocs.io/

## Intra-document anchors

See [below](#target-section) for the target.

(target-section)=
## Target section

This header is reachable from the link in the paragraph above via the
MyST `(label)=` anchor syntax, which produces a docutils target node.

## Link in a list

- [Python](https://www.python.org/) — the language
- [PyPI](https://pypi.org/) — the index
- [docxsphinx on GitHub](https://github.com/mherkazandjian/docxsphinx)

## Link with inline formatting

A paragraph containing [**bolded link text**](https://example.com/),
a [link with `inline code`](https://example.org/), and a
[~~struck link~~](https://example.net/) — mostly to verify the run
formatting inside a hyperlink works correctly once Phase 2.2 lands.
