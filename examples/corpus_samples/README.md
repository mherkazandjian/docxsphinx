# corpus_samples/ — committed fixtures for the roundtrip harness

Tiny, reproducible `.docx` fixtures used by CI / `make roundtrip` when the
private `corpus/` directory is empty. The docs here are generated from our
own Markdown sources (so they're MIT-licensed along with the rest of the
repo) but by **external tools** (pandoc's docx output, LibreOffice, etc.)
— i.e. deliberately not by docxsphinx. The harness compares what these
external tools produce against what docxsphinx produces for the same
input; the divergences are the research output.

## What's here

- `pandoc_md_basic.docx` — pandoc's docx output from
  `examples/md_basic/source/index.md`. Regenerate with:
  ```
  docker compose run --rm dev pandoc -f gfm -t docx \
      -o examples/corpus_samples/pandoc_md_basic.docx \
      examples/md_basic/source/index.md
  ```

Add additional fixtures (other tools, other source flavours) as the
research harness reveals gaps worth studying.
