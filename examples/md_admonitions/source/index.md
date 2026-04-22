# md_admonitions — MyST admonition directives

Admonitions are supported in MyST via the `:::{note}` / `::: {warning}` /
etc. colon-fence syntax. Until Phase 2.1 of the modernization plan
lands, docxsphinx silently drops them — this sample will render as a
near-empty document. Once admonitions are wired up, every block below
should appear as a distinct bordered callout.

## Note

:::{note}
This is a short **note** callout with a little inline `code` inside.
:::

## Warning

:::{warning}
Proceed with caution: the operation you are about to perform cannot
easily be undone.
:::

## Tip

:::{tip}
You can use `make example DIR=md_admonitions` to rebuild this sample
directly in the dev container.
:::

## Caution, important, hint, error, danger, attention

:::{caution}
Moving parts — keep fingers clear.
:::

:::{important}
Read the release notes before upgrading.
:::

:::{hint}
Try `pytest -m unit` for a faster inner-loop.
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

## Custom admonition

:::{admonition} Custom heading
With custom body content.
:::

## Admonition with a multi-paragraph body

:::{note}
Paragraph one inside the admonition.

Paragraph two, separated by a blank line. Admonitions should preserve
internal structure — lists, code, inline formatting, and so on.

- A bullet inside the admonition
- Another bullet

```python
# Even a code block inside the admonition.
print("hello from inside a note")
```
:::
