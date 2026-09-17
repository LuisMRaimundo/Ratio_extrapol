"""Parse individual or block-pasted acoustic tables onto a chromatic grid."""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from typing import Optional

from .notes import Pitch, parse_pitch

_HEADER_NOTE = re.compile(r"^(source\s*)?notes?$|note_sounding|pitch|midi", re.I)
_HEADER_VALUE = re.compile(
    r"cdm|combined density|value|metric|spectral|mass|density|media|number|val$",
    re.I,
)
_HEADER_ORIGIN = re.compile(r"origin|estimate|status|kind|source_type", re.I)
_HEADER_DYNAMIC = re.compile(r"dynamic", re.I)
_HEADER_TECH = re.compile(r"technique|state", re.I)
_HEADER_COLL = re.compile(r"collection|corpus|library", re.I)
_HEADER_INSTR = re.compile(r"instrument", re.I)
_HEADER_COMMENT = re.compile(r"comment|flags?$", re.I)
_HEADER_PI_LOW = re.compile(r"pi95\s*low|pi.?low", re.I)
_HEADER_PI_HIGH = re.compile(r"pi95\s*high|pi.?high", re.I)
_HEADER_STATUS = re.compile(r"reporting_status|cell_status", re.I)
_HEADER_ANCHOR = re.compile(r"anchor_source", re.I)
_HEADER_MIDI = re.compile(r"^midi$", re.I)

_FLOAT_RE = re.compile(
    r"""
    ^[+-]?(?:
        (?:\d+(?:[.,]\d*)?) |
        (?:[.,]\d+)
    )(?:[eE][+-]?\d+)?$
    """,
    re.VERBOSE,
)


@dataclass
class ParsedRow:
    pitch: Optional[Pitch]
    value: Optional[float]
    origin: str = ""
    dynamic: str = ""
    technique: str = ""
    collection: str = ""
    instrument: str = ""
    comment: str = ""
    raw: str = ""
    pi95_low: Optional[float] = None
    pi95_high: Optional[float] = None
    reporting_status: str = ""
    anchor_source: str = ""


@dataclass
class ParseReport:
    rows: list[ParsedRow] = field(default_factory=list)
    format_name: str = "unknown"
    messages: list[str] = field(default_factory=list)
    start_note_needed: bool = False
    value_only_count: int = 0

    @property
    def placed_pairs(self) -> list[ParsedRow]:
        return [r for r in self.rows if r.pitch and r.value is not None]


def parse_number(text: object) -> Optional[float]:
    if text is None:
        return None
    s = str(text).strip()
    if not s or s.lower() in {"nan", "none", "-", ""}:
        return None
    s = s.replace("\u00a0", "").replace(" ", "")
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    elif s.count(".") > 1 and s.count(",") == 1:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _looks_number(text: str) -> bool:
    return parse_number(text) is not None and parse_pitch(text) is None


def _split_lines(block: str) -> list[list[str]]:
    text = (block or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return []
    sample = text.split("\n", 1)[0]
    dialect = ","
    if "\t" in sample:
        dialect = "\t"
    elif ";" in sample and sample.count(";") >= sample.count(","):
        dialect = ";"
    rows: list[list[str]] = []
    reader = csv.reader(io.StringIO(text), delimiter=dialect)
    for raw in reader:
        cells = [c.strip() for c in raw if str(c).strip() != ""]
        if not cells:
            # keep positional empties if the line had content with trailing tabs
            continue
        if len(raw) > 1:
            cells = [c.strip() for c in raw]
        rows.append(cells)
    return rows


def _header_map(cells: list[str]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for i, cell in enumerate(cells):
        if _HEADER_NOTE.match(cell) and "note" not in mapping:
            mapping["note"] = i
        elif _HEADER_VALUE.search(cell) and "value" not in mapping:
            mapping["value"] = i
        elif _HEADER_ORIGIN.search(cell) and "origin" not in mapping:
            mapping["origin"] = i
        elif _HEADER_DYNAMIC.match(cell):
            mapping["dynamic"] = i
        elif _HEADER_TECH.search(cell):
            mapping["technique"] = i
        elif _HEADER_COLL.search(cell):
            mapping["collection"] = i
        elif _HEADER_INSTR.search(cell):
            mapping["instrument"] = i
        elif _HEADER_COMMENT.search(cell) and "comment" not in mapping:
            mapping["comment"] = i
        elif _HEADER_PI_LOW.search(cell):
            mapping["pi95_low"] = i
        elif _HEADER_PI_HIGH.search(cell):
            mapping["pi95_high"] = i
        elif _HEADER_STATUS.search(cell):
            mapping["reporting_status"] = i
        elif _HEADER_ANCHOR.search(cell):
            mapping["anchor_source"] = i
        elif _HEADER_MIDI.match(cell):
            mapping["midi"] = i
    return mapping


def parse_block(block: str, start_note: Optional[str] = None) -> ParseReport:
    report = ParseReport()
    table = _split_lines(block)
    if not table:
        report.messages.append("Nothing to parse.")
        return report

    header_map = _header_map(table[0])
    body = table
    if header_map.get("note") is not None and header_map.get("value") is not None:
        report.format_name = "header_table"
        body = table[1:]
        for raw in body:
            def _get(key: str) -> str:
                idx = header_map.get(key)
                if idx is None or idx >= len(raw):
                    return ""
                return raw[idx]

            pitch = parse_pitch(_get("note"))
            if pitch is None and _get("midi"):
                pitch = parse_pitch("midi" + "".join(ch for ch in _get("midi") if ch.isdigit()))
            value = parse_number(_get("value"))
            report.rows.append(
                ParsedRow(
                    pitch=pitch,
                    value=value,
                    origin=_get("origin"),
                    dynamic=_get("dynamic"),
                    technique=_get("technique"),
                    collection=_get("collection"),
                    instrument=_get("instrument"),
                    comment=_get("comment"),
                    raw=" | ".join(raw),
                    pi95_low=parse_number(_get("pi95_low")),
                    pi95_high=parse_number(_get("pi95_high")),
                    reporting_status=_get("reporting_status"),
                    anchor_source=_get("anchor_source"),
                )
            )
        _sort_and_annotate(report)
        return report

    # Two-token lines: note + value (any order), possibly extra origin.
    pair_rows = 0
    value_only: list[float] = []
    for raw in table:
        tokens = [t for t in raw if t != ""]
        if not tokens:
            continue
        pitches = [parse_pitch(t) for t in tokens]
        numbers = [parse_number(t) for t in tokens]
        pitch = next((p for p in pitches if p), None)
        number_idxs = [i for i, n in enumerate(numbers) if n is not None and pitches[i] is None]
        value = numbers[number_idxs[0]] if number_idxs else None
        origin = ""
        leftovers = [
            tokens[i]
            for i, t in enumerate(tokens)
            if pitches[i] is None and (i not in number_idxs)
        ]
        if leftovers:
            origin = leftovers[0]
        if pitch is None and len(number_idxs) >= 2:
            maybe_midi = numbers[number_idxs[0]]
            maybe_val = numbers[number_idxs[1]]
            if maybe_midi is not None and maybe_val is not None and float(maybe_midi).is_integer() and 0 <= int(maybe_midi) <= 127:
                pitch = parse_pitch(f"midi{int(maybe_midi)}")
                value = maybe_val
                number_idxs = number_idxs[2:]
        if pitch and value is not None:
            pair_rows += 1
            report.rows.append(ParsedRow(pitch=pitch, value=value, origin=origin, raw=" ".join(tokens)))
        elif value is not None and pitch is None and len(tokens) == 1:
            value_only.append(value)
            report.rows.append(ParsedRow(pitch=None, value=value, raw=tokens[0]))
        elif pitch and value is None:
            report.messages.append(f"Note without a number ignored: {pitch.label}")
        else:
            report.messages.append(f"Unreadable line ignored: {' '.join(tokens)[:80]}")

    if pair_rows:
        report.format_name = "note_value_pairs"
        report.rows = [r for r in report.rows if r.pitch and r.value is not None]
        _sort_and_annotate(report)
        return report

    if value_only:
        report.value_only_count = len(value_only)
        start = parse_pitch(start_note) if start_note else None
        if start is None:
            report.format_name = "value_column"
            report.start_note_needed = True
            report.messages.append(
                f"Found {len(value_only)} numbers with no note labels. "
                "Set a start note so they can be laid onto the chromatic scale."
            )
            return report
        report.format_name = "value_column_from_start"
        report.rows = []
        for i, value in enumerate(value_only):
            midi = start.midi + i
            pitch = parse_pitch(f"midi{midi}")
            report.rows.append(ParsedRow(pitch=pitch, value=value, raw=str(value)))
        _sort_and_annotate(report)
        return report

    report.messages.append("Could not detect note/value pairs or a numeric column.")
    return report


def _sort_and_annotate(report: ParseReport) -> None:
    paired = report.placed_pairs
    if not paired:
        return
    original_midis = [r.pitch.midi for r in paired if r.pitch]
    sorted_rows = sorted(paired, key=lambda r: r.pitch.midi if r.pitch else 0)
    if original_midis != [r.pitch.midi for r in sorted_rows if r.pitch]:
        report.messages.append(
            f"Pasted notes were out of chromatic order. {len(sorted_rows)} values "
            "were snapped onto MIDI slots and sorted low → high."
        )
    seen: dict[int, ParsedRow] = {}
    dupes = 0
    for row in sorted_rows:
        midi = row.pitch.midi
        if midi in seen:
            dupes += 1
            report.messages.append(
                f"Duplicate pitch {row.pitch.label} (MIDI {midi}); last pasted value is kept."
            )
        seen[midi] = row
    report.rows = list(seen.values())
    if not report.messages:
        report.messages.append(f"Parsed {len(report.rows)} note/value pairs ({report.format_name}).")
    else:
        report.messages.insert(0, f"Parsed {len(report.rows)} unique pitches ({report.format_name}).")
