# corpus/ — private roundtrip corpus (gitignored)

Drop your own Microsoft-Word-authored `.docx` files in this directory. The
`make roundtrip` target converts each one back to Markdown via pandoc, runs
that Markdown through docxsphinx's forward path, and diffs the original vs
round-trip OOXML. The diff surfaces idiomatic patterns in real Word output
that docxsphinx should match.

This directory is `.gitignored` — nothing under it gets committed. If you
want a reproducible, committable fixture instead, put it in
`../corpus_samples/`.

The aggregate report after a run lives at `reports/` (also gitignored).
