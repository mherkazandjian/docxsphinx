# md_code — fenced code blocks

## Python

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

## Bash

```bash
#!/usr/bin/env bash
set -euo pipefail

for f in *.md; do
    echo "processing $f"
    wc -l "$f"
done
```

## JSON

```json
{
  "project": "docxsphinx",
  "version": "2.0.0",
  "dependencies": ["docutils", "Sphinx", "python-docx", "Pygments"]
}
```

## No language (plain preformatted)

```
                    .-.
                   (o.o)
                    |=|
                   __|__
                 //.=|=.\\
                // .=|=. \\
                \\ .=|=. //
                 \\(_=_)//
                  (:| |:)
                   || ||
                   () ()
                   || ||
                   || ||
                  ==' '==
```

## Inline code

A paragraph with an inline `run_tests()` reference and another
`Document.add_paragraph('hello')` call, followed by prose.
