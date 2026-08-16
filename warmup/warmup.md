<!--
Build-time input for the tectonic bundle warm-up (see the Dockerfile).

Tectonic downloads the TeX support files it needs on demand. This document holds one instance of every
construct the service converts, so that the build pulls the whole set into the image cache and no
conversion at run time depends on the network.

Keep it exhaustive. A construct that is missing here is a file that is missing from the cache.
tests/container/container-structure-test.yaml asserts the result.
-->
---
title: Tectonic bundle warm-up
author: pandoc-service
date: 2026-01-01
header-includes:
  - \usepackage{colortbl}
  - \usepackage{soul}
  - \usepackage[normalem]{ulem}
---

# Heading level one

## Heading level two

### Heading level three

Body text with **bold**, *italic*, ***bold italic***, ~~strikeout~~, `inline code`,
[a link](https://example.com), a footnote[^1], H~2~O and x^2^.

[^1]: Footnote text.

Inline math $E = mc^2$ and display math:

$$\sum_{i=0}^{n} \frac{x_i}{2} \leq \int_0^1 x^2\,dx$$

| Left | Center | Right |
|:-----|:------:|------:|
| 1    |   2    |     3 |
| 4    |   5    |     6 |

: Table caption.

![Image caption](warmup.png)

```python
# Comment: syntax highlighting needs the italic and bold monospace faces.
def warm(engine: str) -> str:
    return f"warm up {engine}"
```

> Block quote.

- Bullet item
- Bullet item

1. Ordered item
2. Ordered item

Term
: Definition.

---
