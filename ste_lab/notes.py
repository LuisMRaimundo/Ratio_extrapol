"""Concert-pitch note grid, enharmonic spelling, and MIDI placement."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

PC_TO_FLAT = ["C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "B"]
PC_TO_SHARP = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]

NAME_TO_PC = {
    "C": 0,
    "C#": 1,
    "DB": 1,
    "D": 2,
    "D#": 3,
    "EB": 3,
    "E": 4,
    "F": 5,
    "F#": 6,
    "GB": 6,
    "G": 7,
    "G#": 8,
    "AB": 8,
    "A": 9,
    "A#": 10,
    "BB": 10,
    "B": 11,
}

# Prefer flats in the working register (matches the viola workbook core).
FLAT_PREFERRED_PC = {1, 3, 6, 8, 10}

_NOTE_RE = re.compile(
    r"""
    ^\s*
    (?P<name>[A-Ga-g])
    (?:
        (?P<acc>\#|b|♯|♭|is|es)
    )?
    \s*
    (?P<oct>-?\d)?
    \s*$
    """,
    re.VERBOSE,
)

_MIDI_RE = re.compile(r"^\s*(?:midi\s*)?(\d{1,3})\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class Pitch:
    midi: int
    label: str
    pitch_class: int
    octave: int

    @property
    def sharp_label(self) -> str:
        return midi_to_label(self.midi, prefer_flats=False)

    @property
    def flat_label(self) -> str:
        return midi_to_label(self.midi, prefer_flats=True)


def midi_to_label(midi: int, prefer_flats: bool = True) -> str:
    pc = midi % 12
    octave = midi // 12 - 1
    names = PC_TO_FLAT if prefer_flats else PC_TO_SHARP
    return f"{names[pc]}{octave}"


def label_to_midi(text: str) -> Optional[int]:
    parsed = parse_pitch(text)
    return parsed.midi if parsed else None


def parse_pitch(text: object) -> Optional[Pitch]:
    if text is None:
        return None
    raw = str(text).strip()
    if not raw or raw.lower() in {"nan", "none", "-"}:
        return None

    raw = (
        raw.replace("♯", "#")
        .replace("♭", "b")
        .replace("\u266f", "#")
        .replace("\u266d", "b")
    )
    raw = raw.replace(" ", "")

    midi_match = _MIDI_RE.match(raw)
    if midi_match and raw.lower().startswith("midi"):
        midi = int(midi_match.group(1))
        if 0 <= midi <= 127:
            return Pitch(midi, midi_to_label(midi), midi % 12, midi // 12 - 1)
        return None

    match = _NOTE_RE.match(raw)
    if not match:
        return None
    name = match.group("name").upper()
    acc = match.group("acc") or ""
    oct_s = match.group("oct")
    if acc.lower() == "is":
        acc = "#"
    elif acc.lower() == "es":
        acc = "b"
    token = (name + acc).upper().replace("#", "#")
    if token not in NAME_TO_PC:
        return None
    pc = NAME_TO_PC[token]
    if oct_s is None:
        return None
    octave = int(oct_s)
    midi = 12 * (octave + 1) + pc
    if not 0 <= midi <= 127:
        return None
    prefer_flats = pc in FLAT_PREFERRED_PC and "#" not in token
    label = raw[0].upper() + raw[1:]
    if prefer_flats and "#" not in token:
        label = midi_to_label(midi, prefer_flats=True)
    return Pitch(midi, label, pc, octave)


def chromatic_range(low_midi: int, high_midi: int, prefer_flats: bool = True) -> list[Pitch]:
    low = max(0, int(low_midi))
    high = min(127, int(high_midi))
    if high < low:
        low, high = high, low
    out = []
    for midi in range(low, high + 1):
        label = midi_to_label(midi, prefer_flats=prefer_flats)
        out.append(Pitch(midi, label, midi % 12, midi // 12 - 1))
    return out


def same_pitch(a: str, b: str) -> bool:
    pa, pb = parse_pitch(a), parse_pitch(b)
    return bool(pa and pb and pa.midi == pb.midi)


def normalize_label(text: str, prefer_flats: bool = True) -> str:
    parsed = parse_pitch(text)
    if not parsed:
        return str(text).strip()
    return midi_to_label(parsed.midi, prefer_flats=prefer_flats)
