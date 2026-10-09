"""Read a whole number an OOXML attribute states, from a template or from pandoc.

``str.isdigit`` accepts characters ``int`` refuses, such as a superscript two, and ``int`` refuses a
string of more than 4300 digits. A template stating either must not fail the conversion, so a value
is read only when it is ASCII digits no longer than a 32-bit OOXML number can be.
"""

# OOXML states its whole numbers in 32 bits: no more than ten digits.
_MAX_DIGITS = 10


def whole_number(value: str | None, *, signed: bool = False) -> int | None:
    """The whole number a value states, a leading minus allowed where signed, else None."""
    if value is None:
        return None
    digits = value[1:] if signed and value.startswith("-") else value
    if not digits.isascii() or not digits.isdigit() or len(digits) > _MAX_DIGITS:
        return None
    return int(value)
