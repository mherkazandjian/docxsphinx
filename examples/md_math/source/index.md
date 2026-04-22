# md_math — LaTeX equations rendered to Word OMML

This example demonstrates docxsphinx's LaTeX math rendering. With pandoc
installed, every equation below lands in the Word document as native
Office Math Markup Language (OMML) — editable in Word's equation
editor, copy-pasteable to other Word documents, and properly rendered
in Word's preview and print views.

Without pandoc, the writer falls back to emitting the raw LaTeX as
monospace text (no crash — just lower fidelity) so a user installing
only the forward pipeline still gets readable output.

## Inline math

Einstein's mass-energy equivalence: $E = mc^2$. The gravitational
constant in natural units is $G = 1$. A quadratic with roots at
$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$.

## Display math

The arithmetic progression sum:

$$
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}
$$

Euler's identity — one of the most celebrated equations in mathematics:

$$
e^{i\pi} + 1 = 0
$$

A piecewise function:

$$
f(x) = \begin{cases}
    x^2 & \text{if } x \geq 0 \\
    -x^2 & \text{if } x < 0
\end{cases}
$$

A matrix:

$$
A = \begin{pmatrix}
    a_{11} & a_{12} & a_{13} \\
    a_{21} & a_{22} & a_{23} \\
    a_{31} & a_{32} & a_{33}
\end{pmatrix}
$$

## AMS environments

The ``amsmath`` MyST extension lets standalone AMS-LaTeX environments
work without surrounding `$$`:

\begin{align}
(a+b)^2 &= a^2 + 2ab + b^2 \\
(a-b)^2 &= a^2 - 2ab + b^2 \\
(a+b)(a-b) &= a^2 - b^2
\end{align}

## Math mixed with prose

The Gaussian integral, $\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}$,
is a cornerstone of probability theory. Its discrete analogue, the sum
$\sum_{k=0}^{\infty} \frac{1}{k!} = e$, defines Euler's number.
