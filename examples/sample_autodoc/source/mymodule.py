"""A tiny demonstration module for the sample_autodoc example.

Provides one class and one module-level function, each documented with
Sphinx-flavoured ``:param:`` / ``:type:`` / ``:returns:`` / ``:raises:``
fields so autodoc produces the full range of node types docxsphinx must
render (``desc_signature``, ``desc_name``, ``desc_parameterlist``,
``desc_parameter``, ``desc_optional``, ``desc_returns``, ``desc_annotation``,
and ``field_list`` / ``field_name`` / ``field_body`` inside ``desc_content``).
"""
from __future__ import annotations


def greet(name: str, greeting: str = 'Hello') -> str:
    """Return a friendly greeting for *name*.

    :param name: The person to greet.
    :type name: str
    :param greeting: Optional salutation prefix.
    :type greeting: str
    :returns: The composed greeting string.
    :rtype: str
    :raises ValueError: If *name* is empty.
    """
    if not name:
        raise ValueError('name must be non-empty')
    return f'{greeting}, {name}!'


class Calculator:
    """A tiny calculator demonstrating class + method autodoc.

    :param initial: Starting value for the accumulator.
    :type initial: float
    """

    def __init__(self, initial: float = 0.0) -> None:
        self.value = initial

    def add(self, x: float, y: float = 0.0) -> float:
        """Add *x* (and optional *y*) to the running value.

        :param x: Primary addend.
        :type x: float
        :param y: Optional secondary addend.
        :type y: float
        :returns: The updated accumulator value.
        :rtype: float
        """
        self.value += x + y
        return self.value

    def reset(self) -> None:
        """Reset the accumulator to zero."""
        self.value = 0.0
