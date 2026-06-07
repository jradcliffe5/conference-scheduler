#!/usr/bin/env python3
"""
schedule_config.py  —  Shared per-conference configuration loader.

Both generate_schedule.py and fill_schedule.py import this so a
single conference.yaml is the only thing you edit per conference.  Load it with:

    from schedule_config import load
    conf = load("conference.yaml")
    conf.NUM_DAYS, conf.SESSION_BLOCKS, ...

The returned Conf dataclass mirrors the constant names the original scripts
used, so the rest of the code reads almost identically (SESSION_BLOCKS becomes
conf.SESSION_BLOCKS, etc.).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

try:
    import yaml
except ImportError:  # pragma: no cover - dependency guard
    raise SystemExit(
        "PyYAML is required to read the conference config.\n"
        "Install it with:  pip install pyyaml"
    )


# ─── Defaults ─────────────────────────────────────────────────────────────────
# Anything not set in the YAML falls back to these (tuned to the EVN sheet layout
# so a minimal YAML still works on an EVN-style workbook).

DEFAULT_COLUMN_ALIASES: dict[str, list[str]] = {
    "id":               ["Abstract ID", "ID"],
    "title":            ["Title / Topic", "Title", "Title/Topic", "Topic"],
    "avg_score":        ["Avg Score", "Average Score", "Score"],
    "suggested_format": ["Suggested Format", "Format"],
    "submitter":        ["Submitter", "Author"],
    "topic_tag":        ["Comments", "Topic", "Tag"],
}

DEFAULT_ABSTRACTS_ALIASES: dict[str, list[str]] = {
    "id":        ["Abstract ID", "ID"],
    "title":     ["Title"],
    "submitter": ["Author", "Submitter"],
}

# suggested-format string (case-insensitive)  →  (talk_type, minutes)
DEFAULT_FORMAT_MAP: dict[str, dict[str, Any]] = {
    "invited talk": {"type": "Invited", "minutes": 30},
    "long oral":    {"type": "Long",    "minutes": 30},
    "short oral":   {"type": "Short",   "minutes": 15},
    "poster":       {"type": "Poster",  "minutes": 2},
}
# talk_type used when a suggested-format value matches nothing in FORMAT_MAP
DEFAULT_FALLBACK_TYPE = "Poster"

REQUIRED_KEYS = ["conf_start", "num_days", "session_blocks"]


@dataclass
class Conf:
    # ── identity / files ──────────────────────────────────────────────────────
    CONF_NAME:       str
    CONF_VENUE:      str
    CONF_START:      date
    NUM_DAYS:        int
    INPUT_ABSTRACTS: str
    SCHEDULE_FILE:   str
    OUTPUT_FILE:     str          # generate_schedule writes this (== SCHEDULE_FILE)

    # ── grid structure ────────────────────────────────────────────────────────
    SESSION_BLOCKS:  list[tuple]
    SLOT_MIN:        int
    BREAKS:          list[tuple]
    FIXED_ENTRIES:   list[tuple]
    SPARKLER_SESSIONS: list[tuple]

    # ── talk durations ────────────────────────────────────────────────────────
    LONG_MIN:        int
    SHORT_MIN:       int
    SPARKLER_MIN:    int

    # ── filling behaviour ─────────────────────────────────────────────────────
    INVITED_IDS:     list[str]
    RANDOM_SEED:     int | None
    STATION_TAG:     str
    STATION_THRESHOLD: float
    STATION_SESSIONS: list[tuple]

    # ── spreadsheet mapping (the genericity layer) ────────────────────────────
    SCORES_SHEET:      str
    ABSTRACTS_SHEET:   str
    COLUMN_ALIASES:    dict[str, list[str]]
    ABSTRACTS_ALIASES: dict[str, list[str]]
    FORMAT_MAP:        dict[str, dict[str, Any]]
    FALLBACK_TYPE:     str
    ID_PATTERN:        str
    CONTRIBUTED_ID_PATTERN: str

    # derived: session_index -> bool, "does a break row follow this session?"
    SESSION_HAS_BREAK_AFTER: dict[int, bool] = field(default_factory=dict)

    # Upstream review-sheet generator settings (see generate_review.py).
    # Kept as a raw sub-mapping so that stage's many knobs don't bloat this class.
    REVIEW: dict[str, Any] = field(default_factory=dict)

    def duration_for_type(self, talk_type: str) -> int:
        return {
            "Invited": self.LONG_MIN,
            "Long":    self.LONG_MIN,
            "Short":   self.SHORT_MIN,
        }.get(talk_type, self.SPARKLER_MIN)


def _as_tuples(rows: Any) -> list[tuple]:
    """YAML lists-of-lists -> list of tuples (so existing index code is unchanged)."""
    if not rows:
        return []
    return [tuple(r) for r in rows]


def _as_date(value: Any) -> date:
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        return date.fromisoformat(value.strip())
    raise ValueError(f"conf_start must be an ISO date string (YYYY-MM-DD), got {value!r}")


def _lower_keys(d: dict | None) -> dict:
    """Lower-case the keys of FORMAT_MAP so matching is case-insensitive."""
    return {str(k).strip().lower(): v for k, v in (d or {}).items()}


def load(path: str) -> Conf:
    with open(path, "r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}

    if not isinstance(raw, dict):
        raise ValueError(f"{path}: top-level YAML must be a mapping of settings")

    missing = [k for k in REQUIRED_KEYS if k not in raw]
    if missing:
        raise ValueError(
            f"{path}: missing required key(s): {', '.join(missing)}.\n"
            f"Required: {', '.join(REQUIRED_KEYS)}"
        )

    session_blocks = _as_tuples(raw["session_blocks"])
    breaks         = _as_tuples(raw.get("breaks", []))

    schedule_file = raw.get("schedule_file", "schedule.xlsx")
    output_file   = raw.get("output_file", schedule_file)

    # Merge user column aliases over defaults so partial overrides work.
    col_aliases = dict(DEFAULT_COLUMN_ALIASES)
    col_aliases.update(raw.get("column_aliases", {}) or {})
    abs_aliases = dict(DEFAULT_ABSTRACTS_ALIASES)
    abs_aliases.update(raw.get("abstracts_aliases", {}) or {})

    fmt_map = _lower_keys(DEFAULT_FORMAT_MAP)
    fmt_map.update(_lower_keys(raw.get("format_map")))

    break_after = {b[1]: True for b in breaks}
    session_has_break = {si: break_after.get(si, False)
                         for si in range(len(session_blocks))}

    return Conf(
        CONF_NAME       = raw.get("conf_name", "Conference"),
        CONF_VENUE      = raw.get("conf_venue", ""),
        CONF_START      = _as_date(raw["conf_start"]),
        NUM_DAYS        = int(raw["num_days"]),
        INPUT_ABSTRACTS = raw.get("input_abstracts", "abstract_reviews.xlsx"),
        SCHEDULE_FILE   = schedule_file,
        OUTPUT_FILE     = output_file,

        SESSION_BLOCKS  = session_blocks,
        SLOT_MIN        = int(raw.get("slot_min", 15)),
        BREAKS          = breaks,
        FIXED_ENTRIES   = _as_tuples(raw.get("fixed_entries", [])),
        SPARKLER_SESSIONS = _as_tuples(raw.get("sparkler_sessions", [])),

        LONG_MIN        = int(raw.get("long_min", 30)),
        SHORT_MIN       = int(raw.get("short_min", 15)),
        SPARKLER_MIN    = int(raw.get("sparkler_min", 2)),

        INVITED_IDS     = list(raw.get("invited_ids", []) or []),
        RANDOM_SEED     = raw.get("random_seed", 42),
        STATION_TAG     = raw.get("station_tag", ""),
        STATION_THRESHOLD = float(raw.get("station_threshold", 0.0)),
        STATION_SESSIONS = _as_tuples(raw.get("station_sessions", [])),

        SCORES_SHEET      = raw.get("scores_sheet", "Scores ranked"),
        ABSTRACTS_SHEET   = raw.get("abstracts_sheet", "Abstracts"),
        COLUMN_ALIASES    = {k: list(v) for k, v in col_aliases.items()},
        ABSTRACTS_ALIASES = {k: list(v) for k, v in abs_aliases.items()},
        FORMAT_MAP        = fmt_map,
        FALLBACK_TYPE     = raw.get("fallback_type", DEFAULT_FALLBACK_TYPE),
        ID_PATTERN        = raw.get("id_pattern", r"^[A-Z]+\d+$"),
        CONTRIBUTED_ID_PATTERN = raw.get("contributed_id_pattern", r"^A\d+$"),

        SESSION_HAS_BREAK_AFTER = session_has_break,

        REVIEW = raw.get("review", {}) or {},
    )
