"""Controlled vocabularies for instruments, techniques, collections, origins."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


TECHNIQUES = [
    "ordinario",
    "arco",
    "con sordino",
    "sul tasto",
    "sul ponticello",
    "harmonics",
    "au talon",
    "pizzicato",
    "tremolo",
    "col legno",
    "mute",
    "flutter",
    "custom",
]

DYNAMICS = ["pp", "p", "mp", "mf", "f", "ff"]

COLLECTIONS = ["IOWA", "ORCH", "Orchidea", "Philharmonia", "SHARC", "Studio", "Custom"]

ORIGINS = [
    "measured",
    "modelled_IOWA_anchored",
    "modelled_ORCHIDEA_anchored",
    "modelled_on_extrapolated_anchor",
    "generated",
    "interpolated",
    "extrapolated_polynomial",
    "extrapolated_ridge",
    "extrapolated_pchip",
    "extrapolated_hold",
    "technique_transfer",
    "family_transfer",
    "orchidea_family_transfer_estimate",
    "combined_estimate",
    "collection_anchored",
    "modelled",
    "manual",
]

REPORTING_STATUSES = ["measured_range", "above_measured_ceiling"]
DEFAULT_MEASURED_CEILING_MIDI = 100

PITCH_BASES = ["sounding_concert", "written"]

FAMILIES = {
    "bowed_strings": "Bowed strings",
    "clarinets": "Clarinet family",
    "flutes": "Flute family",
    "oboes": "Oboe family",
    "bassoons": "Bassoon family",
    "saxophones": "Saxophone family",
    "trumpets": "Trumpet family",
    "trombones": "Trombone family",
    "horns": "Horn family",
    "other": "Other / custom",
}


@dataclass(frozen=True)
class InstrumentSpec:
    instrument_id: str
    display_name: str
    family: str
    sounding_low: int
    sounding_high: int
    comfortable_low: int
    comfortable_high: int
    harmonics_sounding_low: Optional[int]
    transposition: int = 0
    aliases: tuple[str, ...] = ()
    notes: str = ""


INSTRUMENTS: dict[str, InstrumentSpec] = {}


def _add(spec: InstrumentSpec) -> None:
    INSTRUMENTS[spec.instrument_id] = spec
    for alias in spec.aliases:
        INSTRUMENTS[alias] = spec


_add(InstrumentSpec("violin", "Violin", "bowed_strings", 55, 93, 55, 88, 67, 0, ("vl",), "Open G3. First practical natural harmonic sounding ~G4."))
_add(InstrumentSpec("viola", "Viola", "bowed_strings", 48, 88, 50, 81, 60, 0, ("vla",), "Open C3. First practical natural harmonic sounding ~C4. Chromatic harmonic maps usually start near C5."))
_add(InstrumentSpec("cello", "Cello", "bowed_strings", 36, 76, 36, 69, 48, 0, ("vc", "violoncello"), "Open C2. First practical natural harmonic sounding ~C3."))
_add(InstrumentSpec("double_bass", "Double bass", "bowed_strings", 28, 67, 28, 55, 40, 0, ("bass", "cb", "contrabass"), "Orchestra tuning; extension systems vary."))
_add(InstrumentSpec("clarinet", "Clarinet in Bb", "clarinets", 50, 94, 50, 89, None, -2, ("clarinet_bb", "cl"), "Sounding range. Written clarinet sounds a major second lower."))
_add(InstrumentSpec("a_clarinet", "Clarinet in A", "clarinets", 49, 93, 49, 88, None, -3, ("clarinet_a",), "Sounding range."))
_add(InstrumentSpec("eb_clarinet", "Clarinet in Eb", "clarinets", 55, 99, 55, 94, None, 3, ("clarinet_eb",), "Sounding range."))
_add(InstrumentSpec("bass_clarinet", "Bass clarinet in Bb", "clarinets", 34, 77, 37, 70, None, -14, ("bcl", "bass_cl"), "Treble-clef convention: written sounds a major ninth lower. Confirm local notation."))
_add(InstrumentSpec("contrabass_clarinet", "Contrabass clarinet", "clarinets", 22, 65, 26, 58, None, -26, ("cbcl",), "Family transfer only with measured overlap. Range conventions vary."))
_add(InstrumentSpec("flute", "Flute", "flutes", 59, 96, 60, 91, None, 0, ("fl",), "C concert flute."))
_add(InstrumentSpec("piccolo", "Piccolo", "flutes", 74, 108, 74, 103, None, 12, ("picc",), "Sounds an octave above written."))
_add(InstrumentSpec("alto_flute", "Alto flute in G", "flutes", 55, 91, 55, 86, None, -5, ("afl",), "Sounds a perfect fourth below written."))
_add(InstrumentSpec("bass_flute", "Bass flute", "flutes", 48, 84, 48, 79, None, -12, ("bfl",), "Sounds an octave below written."))
_add(InstrumentSpec("oboe", "Oboe", "oboes", 58, 91, 58, 86, None, 0, ("ob",)))
_add(InstrumentSpec("english_horn", "English horn", "oboes", 52, 84, 52, 79, None, -7, ("cor_anglais", "eh"), "Sounds a perfect fifth below written."))
_add(InstrumentSpec("bassoon", "Bassoon", "bassoons", 34, 77, 34, 72, None, 0, ("bn", "fg")))
_add(InstrumentSpec("contrabassoon", "Contrabassoon", "bassoons", 22, 65, 22, 53, None, -12, ("cbn",), "Sounds an octave below written."))
_add(InstrumentSpec("soprano_sax", "Soprano saxophone", "saxophones", 56, 89, 56, 84, None, -2, ("ssax",)))
_add(InstrumentSpec("alto_sax", "Alto saxophone", "saxophones", 49, 81, 49, 77, None, -9, ("asax",)))
_add(InstrumentSpec("tenor_sax", "Tenor saxophone", "saxophones", 44, 76, 44, 72, None, -14, ("tsax",)))
_add(InstrumentSpec("baritone_sax", "Baritone saxophone", "saxophones", 36, 69, 36, 65, None, -21, ("bsax",)))
_add(InstrumentSpec("trumpet", "Trumpet in Bb", "trumpets", 58, 94, 58, 89, None, -2, ("tpt",)))
_add(InstrumentSpec("trombone", "Tenor trombone", "trombones", 40, 77, 40, 72, None, 0, ("tbn",)))
_add(InstrumentSpec("bass_trombone", "Bass trombone", "trombones", 34, 72, 34, 65, None, 0, ("btbn",)))
_add(InstrumentSpec("horn", "Horn in F", "horns", 41, 77, 41, 72, None, -7, ("hn",)))
_add(InstrumentSpec("custom", "Custom instrument", "other", 24, 108, 36, 84, None, 0, ()))


def instrument_choices(family: Optional[str] = None) -> list[InstrumentSpec]:
    seen: set[str] = set()
    out: list[InstrumentSpec] = []
    for spec in INSTRUMENTS.values():
        if spec.instrument_id in seen:
            continue
        seen.add(spec.instrument_id)
        if family and spec.family != family:
            continue
        out.append(spec)
    return sorted(out, key=lambda s: (s.family, s.sounding_low, s.display_name))


def resolve_instrument(name: str) -> Optional[InstrumentSpec]:
    key = (name or "").strip().lower().replace(" ", "_")
    if key in INSTRUMENTS:
        return INSTRUMENTS[key]
    for spec in instrument_choices():
        if spec.display_name.lower() == (name or "").strip().lower():
            return spec
    return None


def same_family(a: str, b: str) -> bool:
    sa, sb = resolve_instrument(a), resolve_instrument(b)
    return bool(sa and sb and sa.family == sb.family and sa.family != "other")


# Host woodwind → same-family donor used for Orchidea transfer (never the host).
FAMILY_TRANSFER_DONOR = {
    "bass_clarinet": "clarinet",
    "contrabass_clarinet": "clarinet",
    "a_clarinet": "clarinet",
    "eb_clarinet": "clarinet",
    "piccolo": "flute",
    "alto_flute": "flute",
    "bass_flute": "flute",
    "english_horn": "oboe",
    "contrabassoon": "bassoon",
}


def family_donor_instrument(name: str) -> str | None:
    """Sibling used as the Orchidea / Iowa-sibling donor. None when the instrument is the family principal."""
    spec = resolve_instrument(name)
    if not spec:
        return None
    donor = FAMILY_TRANSFER_DONOR.get(spec.instrument_id)
    if donor and donor != spec.instrument_id:
        return donor
    return None


# Orchestral colour groups used on Media pp / mf / ff columns.
ORCHESTRAL_GROUP_FAMILIES = {
    "strings": frozenset({"bowed_strings"}),
    "brass": frozenset({"trumpets", "trombones", "horns"}),
    "woodwinds": frozenset({"clarinets", "flutes", "oboes", "bassoons", "saxophones"}),
}


def orchestral_group(name: str) -> str:
    """Map an instrument or family id to strings / brass / woodwinds / other."""
    raw = (name or "").strip().lower().replace(" ", "_")
    if raw in {"strings", "brass", "woodwinds", "other"}:
        return raw
    if raw == "woodwind":
        return "woodwinds"
    spec = resolve_instrument(name)
    family = spec.family if spec else raw
    for group, families in ORCHESTRAL_GROUP_FAMILIES.items():
        if family in families:
            return group
    return "other"


ORIGIN_RANK = {
    "measured": 0,
    "collection_anchored": 1,
    "interpolated": 2,
    "technique_transfer": 3,
    "family_transfer": 4,
    "orchidea_family_transfer_estimate": 4,
    "combined_estimate": 6,
    "modelled_iowa_anchored": 5,
    "modelled_orchidea_anchored": 5,
    "modelled": 5,
    "generated": 6,
    "extrapolated_ridge": 7,
    "extrapolated_pchip": 7,
    "extrapolated_hold": 7,
    "extrapolated_polynomial": 8,
    "modelled_on_extrapolated_anchor": 8,
    "manual": 9,
}


def origin_rank(origin: str) -> int:
    return ORIGIN_RANK.get((origin or "").strip().lower(), 50)


def normalize_origin(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return "manual"
    low = s.lower().replace(" ", "_")
    aliases = {
        "iowa_anchored": "modelled_IOWA_anchored",
        "modelled_iowa_anchored": "modelled_IOWA_anchored",
        "orchidea_anchored": "modelled_ORCHIDEA_anchored",
        "modelled_orchidea_anchored": "modelled_ORCHIDEA_anchored",
        "modelled_on_extrapolated_anchor": "modelled_on_extrapolated_anchor",
        "polynomial_wide": "extrapolated_polynomial",
        "polynomial_edge": "extrapolated_polynomial",
        "local_ridge_edge": "extrapolated_ridge",
        "local_ridge_wide": "extrapolated_ridge",
        "local_isolated": "extrapolated_ridge",
        "extrapolated_ridge": "extrapolated_ridge",
        "extrapolated_pchip": "extrapolated_pchip",
        "extrapolated_hold": "extrapolated_hold",
        "polynomial_regression_3": "extrapolated_polynomial",
    }
    if s in ORIGINS:
        return s
    if low in aliases:
        return aliases[low]
    for item in ORIGINS:
        if item.lower() == low:
            return item
    return s if s else "manual"


GENERATED_ORIGINS = {
    "generated",
    "interpolated",
    "extrapolated_polynomial",
    "extrapolated_ridge",
    "extrapolated_pchip",
    "extrapolated_hold",
    "technique_transfer",
    "family_transfer",
    "orchidea_family_transfer_estimate",
    "combined_estimate",
    "collection_anchored",
    "modelled",
    "modelled_iowa_anchored",
    "modelled_orchidea_anchored",
    "modelled_on_extrapolated_anchor",
}
