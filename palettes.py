"""Named colour palettes and semantic colours for think-cell charts.

Ports the McKinsey Vizro design system's colour values into this generator so
decks can use a consistent, colourblind-safe palette and *meaning-driven*
fills (positive / negative / neutral) instead of ad-hoc hex codes scattered
through each build script.

Two things are exposed:

* **Named palettes** -- ordered accent sequences for ``set_deck_branding`` /
  ``apply_branding`` (they map to PowerPoint theme accent1..accent6).
* **Semantic colours** -- a fixed map (positive/negative/neutral/info/warning)
  plus :func:`semantic_fill`, which turns a *number* into the right fill so a
  waterfall or a delta series colours itself (gains teal, losses red).

All hex values are 6-digit ``#RRGGBB`` strings. The Vizro qualitative sequence
and semantic colours are validated by Vizro to be colourblind-safe.
"""
from __future__ import annotations

import re
from typing import Any

# Vizro qualitative sequence (themes/_colors.py), in order. Used when a deck
# wants the McKinsey-flavoured look rather than the Big-Four default.
VIZRO_QUALITATIVE: list[str] = [
    "#097DFE",  # blue
    "#6F39E3",  # dark purple
    "#05D0F0",  # turquoise
    "#0F766E",  # dark green
    "#8C8DE9",  # light purple
    "#11B883",  # light green
    "#E77EC2",  # light pink
    "#C84189",  # dark pink
    "#C0CA33",  # yellow
    "#3E495B",  # gray
]

# Big-Four-inspired default (kept in sync with branding.DEFAULT_ACCENTS so a
# palette name resolves to the same six colours the template ships with).
BIG_FOUR: list[str] = [
    "#00338D",  # KPMG primary blue
    "#0091DA",  # KPMG light blue
    "#86BC25",  # Deloitte green
    "#E88D14",  # PwC orange
    "#797878",  # EY gray
    "#FFDB00",  # EY yellow
]

# Khyber Pakhtunkhwa TA-team deck palette, matching "template TA team.pptx"'s
# theme (deep teal / teal / mint / gold / amber / deep red).
KP_TA_TEAM: list[str] = [
    "#004B53",  # deep teal
    "#00717D",  # teal
    "#8BD8BA",  # mint
    "#FFDF9F",  # gold
    "#E29700",  # amber
    "#8C0A23",  # deep red
]

# Semantic colours (Vizro). These carry meaning and must stay fixed regardless
# of the chosen accent palette: green=good, red=bad, gray=baseline.
SEMANTIC: dict[str, str] = {
    "positive": "#00B5A9",  # success / gain
    "negative": "#EA5748",  # error / loss
    "neutral": "#3E495B",   # baseline / total
    "info": "#097DFE",      # highlight
    "warning": "#FFC107",   # caution
}

# Registry of named palettes available to the deck builder. ``vizro`` is the
# qualitative sequence trimmed to PowerPoint's six accent slots.
PALETTES: dict[str, list[str]] = {
    "vizro": VIZRO_QUALITATIVE[:6],
    "vizro_full": VIZRO_QUALITATIVE,
    "big_four": BIG_FOUR,
    "kp_ta_team": KP_TA_TEAM,
}

_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def is_hex(value: Any) -> bool:
    """True for a well-formed ``#RRGGBB`` colour string."""
    return isinstance(value, str) and bool(_HEX_RE.match(value))


def get_palette(name: str) -> list[str]:
    """Resolve a palette name to its ordered list of ``#RRGGBB`` accents.

    Args:
        name: A key in :data:`PALETTES` (case-insensitive). Unknown names raise
            ``KeyError`` so a typo fails loudly rather than silently picking a
            default.

    Returns:
        A new list of hex colour strings (safe to mutate).
    """
    key = name.strip().lower() if isinstance(name, str) else ""
    if key not in PALETTES:
        available = ", ".join(sorted(PALETTES))
        raise KeyError(f"unknown palette {name!r}; available: {available}")
    return list(PALETTES[key])


def semantic_fill(value: Any, *, reverse: bool = False) -> str:
    """Return the semantic fill colour for a signed numeric ``value``.

    A positive value reads as a gain (teal), a negative as a loss (red), and
    exactly zero as the neutral baseline (gray). Non-numeric or non-finite
    inputs fall back to neutral so a bad value never crashes a build.

    Args:
        value: The number whose sign drives the colour.
        reverse: Flip the gain/loss mapping. Use this when *down is good* --
            e.g. cost, churn, or defect metrics -- so a decrease shows teal.

    Returns:
        One of the semantic hex colours.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return SEMANTIC["neutral"]
    if value != value or value in (float("inf"), float("-inf")):  # NaN / Inf
        return SEMANTIC["neutral"]
    if value == 0:
        return SEMANTIC["neutral"]
    good = value > 0
    if reverse:
        good = not good
    return SEMANTIC["positive"] if good else SEMANTIC["negative"]


def list_palettes() -> dict[str, Any]:
    """Describe the available palettes and semantic colours for a tool reply."""
    return {
        "palettes": {name: list(colors) for name, colors in PALETTES.items()},
        "semantic": dict(SEMANTIC),
        "default": "big_four",
    }
