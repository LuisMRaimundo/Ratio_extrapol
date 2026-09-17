# Ratio_extrapol — mathematical reference

**Current project name:** **Ratio_extrapol** (GitHub: [LuisMRaimundo/Ratio_extrapol](https://github.com/LuisMRaimundo/Ratio_extrapol); local folder spelling `Ratio_Extrapol`).  
**Status:** production documentation of this Ratio_extrapol working tree on branch `fix/ratio-export-consistency`, starting from `42a4df1f662d5a2f5156d608305bbaff641ce1fb`. Engineering fixes below change export consistency, provenance, configuration validation, and Fill labels. They do **not** establish physical or acoustic validity. Historical research workbooks were **not** regenerated.  
**Implementation identifiers (unchanged in source):** Python package `ste_lab` reports version 1.3.3 and still prints “STE Lab” in the GUI, workbook metadata, and some comments. Those strings are **legacy self-names inside the code**, not the current repository name.  
**Language:** English. StackEdit-compatible Markdown + LaTeX (`$...$` inline, `$$...$$` display). No custom macros.  
**Audit date:** 2026-09-17. Correction pass: 2026-09-17.

This document extracts **project-defined** mathematics from first-party Python and from spreadsheet formulas this Ratio_extrapol repository writes. External-library internals are **not** reproduced; only the call site, arguments, and surrounding project math are recorded.

**This document does not certify scientific validity.** The existence of tests, comments, or exported methodology text is not evidence that those tests were executed in this audit, nor that the implemented estimators are physically or statistically adequate.

Do **not** treat this Ratio_extrapol repository as identical to Extrapol_data, the separate Strings Techniques Extrapolation (STE) product, Dynamics_extrapol, or an older CDM Excel model. Shared vocabulary (F-061, Media, $L$, IOWA/ORCH) and leftover `ste_lab` / “STE Lab” identifiers do not imply shared formulae.

---

## Source provenance

| Item | Value |
|------|--------|
| Current name | Ratio_extrapol (GitHub); local folder `Ratio_Extrapol` |
| Legacy name in source | STE Lab / package `ste_lab` (not renamed in this documentation pass) |
| Local repository root | `C:\Users\lmr20\Desktop\Código extrapolação\Ratio_Extrapol` |
| Local branch | `fix/ratio-export-consistency` (created from `main`) |
| Local HEAD | `42a4df1f662d5a2f5156d608305bbaff641ce1fb` (uncommitted correction working tree) |
| Uncommitted source changes at audit start | none (working tree clean). Correction pass adds `ste_lab/media_policy.py` and focused export/Fill/config edits. Historical analyses were not rewritten. |
| GitHub | https://github.com/LuisMRaimundo/Ratio_extrapol |
| GitHub default branch | `main` |
| GitHub HEAD (read-only: `gh api` + `git ls-remote`; **no fetch** into the local repo) | `42a4df1f662d5a2f5156d608305bbaff641ce1fb` |
| Local vs GitHub | **identical** at documentation time |
| Declared dependencies | `numpy>=1.24`, `scipy>=1.11`, `pandas>=2.0`, `openpyxl>=3.1` ([requirements.txt](../requirements.txt)) |
| Environment used only for version strings and SciPy docstring (project modules were not imported) | Python 3.10.11; `numpy 2.2.6`, `scipy 1.13.1`, `pandas 2.3.3`, `openpyxl 3.1.5` |
| Logarithm in first-party code | natural log: `numpy.log` / `math.log` and `numpy.exp` / `math.exp` — **not** $\log_{10}$, **not** dB, **not** SPL |
| Destination document | [docs/Ratio_Extrapol_math_formula.md](Ratio_Extrapol_math_formula.md); project README [../README.md](../README.md) |

SHA-256 hashes of documented files are in [Appendix B](#appendix-b--source-file-hashes). Recompute if those files change.

Comparison limitation: GitHub was inspected by remote SHA only. A temporary clone was unnecessary because the SHAs matched. This document describes the **local working tree**, which happened to equal GitHub.

---

## Scope, exclusions, and reading guide

### In scope

- First-party Python under the Ratio_extrapol repository root, including the `ste_lab/` package (legacy directory name), batch runners, preparatory builders, and tests as **verification evidence**.
- `calibration.yaml` defaults that change numerical behaviour.
- Spreadsheet formulae **written by** this code (Zenodo `AcousticTable` / `Summary` Excel formulas).
- Catalogue MIDI constants that bound grids, floors, and QA.

### Out of scope (explicit exclusions)

- Routine counters, GUI geometry, progress text, and file-management arithmetic that do not change scientific results (`media_format.py` colour fills; Tk widget layout; path discovery except where it selects which numbers enter a curve).
- Internal mathematics of NumPy, SciPy, pandas, openpyxl, or Excel (`GEOMEAN`, `SLOPE`, `PCHIP`, `polyfit`, …).
- The numerical definition of **F-061 / spectral_mass / Combined Density Metric**. This repository **loads** that number from compiled workbooks; it does not compute it from audio.
- Prediction intervals (`pi95_low`, `pi95_high`): stored and exported if present on paste/import; **no generator** exists here.
- Frequency in hertz, equal-tempered $440\,\mathrm{Hz}$ conversion, and dB/SPL. None are implemented.
- Gaussian processes, Bayesian posteriors, and hierarchical random-effects models. None are implemented.
- Notebooks: **none** in this repository.
- Binary `ste_lab/Viola_Zenodo_collections_harmonics.xlsx`: reference/data file, not a formula generator.

### How to read an entry

Each **M-###** is a project-defined operation. Each **L-###** is a delegated library/spreadsheet call. Identifiers are stable for this document series.

For every M-entry: **A** status, **B** location, **C** verbatim snippet, **D** LaTeX, **E** symbols, **F** layman, **G** specialist, **H** conditions / downstream, **I** tests.

---

## Implemented processing flow

Two production families share the same $L=\ln(\mathrm{target}/\mathrm{source})$ kernel but **not** the same Media/calibration policy.

```text
compiled F-061 workbooks (external)
        │
        ▼
load_spectral_mass / Paste_arco  ──►  measured curves (MIDI → value > 0)
        │
        ├─ strings: Anchors_all pairs → technique_transfer(log_ratio)
        │           recorded Orchidea effect cells kept as measured (no L)
        │           IOWA effect = IOWA_ordinario × exp(clip(L, ±3))
        │           Media = arithmetic mean of IOWA+ORCH peers
        │
        └─ woodwinds: two_log_ratio on sibling clarinet (default)
                      ORCH_target = ORCH_sibling × exp(clip(L_instr, ±3))
                      combination default empirical_only (measurement wins)
                      Media from media_policy.resolve_iowa_orch_pair
                      Final_Calibration + QA distances
        │
        ▼
Ratio_extrapol workbook (`*_STE_*.xlsx` filename pattern) + Zenodo collections workbook
```

Default policy file ([calibration.yaml](../calibration.yaml)): `transfer.method = two_log_ratio`, `combination.method = empirical_only`, validation `mae` / `min_n = 5`, accept $\le 3$ semitones, review $\le 6$ semitones, long extrapolation off, `allow_review_required = false`.

---

## Project formulae

### Stage A — Pitch, catalogue, and grids

### M-001 — Scientific-pitch MIDI encoding (production)

**A. Status:** production. Used whenever a note name is parsed.

**B. Source:** [ste_lab/notes.py](../ste_lab/notes.py), `parse_pitch`, lines 115–119.

**C. Snippet:**

```python
    pc = NAME_TO_PC[token]
    if oct_s is None:
        return None
    octave = int(oct_s)
    midi = 12 * (octave + 1) + pc
```

**D. LaTeX:**

$$\mathrm{MIDI} = 12\,(\mathrm{octave}+1)+\mathrm{pc},\qquad \mathrm{pc}\in\{0,\ldots,11\}$$

accepted only if $0\le \mathrm{MIDI}\le 127$. Bare integers are **values**, not MIDI, unless the token is `midiN` (lines 95–100).

**E. Symbols:** $\mathrm{pc}$ pitch class from `NAME_TO_PC` (C $=0$, …, B $=11$; enharmonics share a class). $\mathrm{octave}$ is the scientific octave in the token (`C4` → 4). Units: MIDI note number (dimensionless integer). Domain: 0–127.

**F. Layman:** The code turns names like `C4` or `Db5` into the standard MIDI number (middle C $=60$).

**G. Specialist:** This is the MIDI Tuning Standard / scientific pitch mapping, not a frequency calculation. Accidentals `is`/`es` are rewritten to `#`/`b`. No transposition is applied here (see M-004).

**H. Downstream:** every ingest, anchor line, and export note label. Wrong octave is a one-octave slip in every later ratio.

**I. Tests:** `test_run_ste_from_preparatory.py` uses `parse_pitch` indirectly; no isolated unit test of the $12(o+1)+\mathrm{pc}$ identity was identified.

---

### M-002 — MIDI to printed label (production)

**A. Status:** production.

**B. Source:** [ste_lab/notes.py](../ste_lab/notes.py), `midi_to_label`, lines 68–72.

**C. Snippet:**

```python
def midi_to_label(midi: int, prefer_flats: bool = True) -> str:
    pc = midi % 12
    octave = midi // 12 - 1
    names = PC_TO_FLAT if prefer_flats else PC_TO_SHARP
    return f"{names[pc]}{octave}"
```

**D. LaTeX:**

$$\mathrm{pc}=\mathrm{MIDI}\bmod 12,\qquad \mathrm{octave}=\lfloor\mathrm{MIDI}/12\rfloor-1$$

Default spelling prefers flats for black keys (`FLAT_PREFERRED_PC`).

**E. Symbols:** integer MIDI in; string label out. Inverse of M-001 for the default spelling.

**F. Layman:** MIDI 60 is printed `C4`; MIDI 61 is `Db4` unless sharps are requested.

**G. Specialist:** Integer division is toward $-\infty$ for negative MIDI, but the accepted domain is $0\ldots127$. Joining tables on the printed string can break on `A#` vs `Bb` (QA flag `enharmonic_spelling`).

**H. Downstream:** every exported note column.

**I. Tests:** No direct test identified.

---

### M-003 — Chromatic grid clamp (production)

**A. Status:** production.

**B. Source:** [ste_lab/notes.py](../ste_lab/notes.py), `chromatic_range`, lines 129–137.

**C. Snippet:**

```python
    low = max(0, int(low_midi))
    high = min(127, int(high_midi))
    if high < low:
        low, high = high, low
    out = []
    for midi in range(low, high + 1):
```

**D. LaTeX:**

$$\{\,m\in\mathbb{Z}: \max(0,\ell)\le m\le \min(127,h)\,\}$$

**E. Symbols:** $\ell,h$ requested MIDI bounds.

**F. Layman:** Build every chromatic step between two notes, staying inside the piano MIDI range.

**G. Specialist:** Inclusive integer range. No frequency interpolation.

**H. Downstream:** `Layer.ensure_grid` in the GUI.

**I. Tests:** No direct test identified.

---

### M-004 — Catalogue sounding range and unused transposition (production constants / unused transform)

**A. Status:** production constants for range checks and layer defaults. The `transposition` field is **stored** and exported; it is **not** added to MIDI in any calculation found.

**B. Source:** [ste_lab/catalog.py](../ste_lab/catalog.py), `InstrumentSpec` and `_add(...)` entries, lines 66–115; copy onto layers in [ste_lab/session.py](../ste_lab/session.py) lines 173–190.

**C. Snippet (representative):**

```python
_add(InstrumentSpec("clarinet", "Clarinet in Bb", "clarinets", 50, 94, 50, 89, None, -2, ("clarinet_bb", "cl"), "Sounding range. Written clarinet sounds a major second lower."))
_add(InstrumentSpec("bass_clarinet", "Bass clarinet in Bb", "clarinets", 34, 77, 37, 70, None, -14, ("bcl", "bass_cl"), "Treble-clef convention: written sounds a major ninth lower. Confirm local notation."))
```

**D. LaTeX:** catalogue interval $[\mathrm{sounding\_low},\mathrm{sounding\_high}]$ in concert MIDI. Transposition $t$ (written minus sounding, in semitones, as documented in the notes strings) satisfies **no** implemented identity $m_{\mathrm{written}}=m_{\mathrm{sounding}}-t$ in this codebase.

**E. Symbols:** MIDI integers. `comfortable_*` unused in mathematics. `harmonics_sounding_low` used in M-005 and QA.

**F. Layman:** Each instrument has a stored concert-pitch range. The code does not convert written parts to sounding MIDI using the transposition number.

**G. Specialist:** Comments claim sounding ranges. If a paste is written-pitch, the catalogue will not correct it. Family-donor map `FAMILY_TRANSFER_DONOR` (lines 147–157) is a discrete lookup, not a formula.

**H. Downstream:** layer `range_low`/`range_high`; family transfer clipping; QA `outside_instrument_range`.

**I. Tests:** `test_family_donor_map`; `test_c_notes_beyond_catalog_sounding_high`.

---

### M-005 — Harmonic-layer MIDI headroom (production)

**A. Status:** production when `technique == "harmonics"`.

**B. Source:** [ste_lab/session.py](../ste_lab/session.py), `layer_from_identity`, lines 177–180. Same rule in [ste_lab/transfer.py](../ste_lab/transfer.py) `family_transfer` lines 176–178.

**C. Snippet:**

```python
    if technique == "harmonics" and spec and spec.harmonics_sounding_low:
        low = spec.harmonics_sounding_low
        high = min(108, spec.sounding_high + 24)
```

**D. LaTeX:**

$$\ell \leftarrow \ell_{\mathrm{harm}},\qquad h\leftarrow \min(108,\,h_{\mathrm{stopped}}+24)$$

**E. Symbols:** $\ell_{\mathrm{harm}}$ catalogue first practical sounding harmonic; $h_{\mathrm{stopped}}$ ordinary sounding high; $+24$ is two octaves in semitones; $108$ is MIDI C8.

**F. Layman:** Harmonic maps are allowed to go two octaves above the stopped top, but not above C8.

**G. Specialist:** Heuristic headroom, not a physical harmonic series. Does not compute $f,2f,3f$.

**H. Downstream:** empty-slot fill and family maps on harmonic layers.

**I. Tests:** No direct test identified.

---

### M-006 — Harmonic floor (drop notes below a MIDI threshold) (production)

**A. Status:** production on string harmonic pipelines. Woodwind rebuilds skip string harmonic headers.

**B. Source:** [run_ste_from_preparatory.py](../run_ste_from_preparatory.py) lines 42–48 and 301–304; [build_ste_preparatory.py](../build_ste_preparatory.py) lines 31–36 and 110–113; instrument CLIs (`HARM_LO` in cello/viola/bass/violin4 batches).

**C. Snippet:**

```python
HARM_FLOOR = {
    "violin": 79,
    "viola": 72,
    "cello": 60,
    "double_bass": 52,
}
```

```python
def harm_floor(instrument: str) -> int:
    spec = resolve_instrument(instrument)
    key = spec.instrument_id if spec else (instrument or "viola").strip().lower().replace(" ", "_")
    return HARM_FLOOR.get(key, HARM_LO)
```

`HARM_LO = 72` is the fallback (viola C5). `build_ste_preparatory.harm_floor` returns `None` if the instrument is not in the table (no fallback).

**D. LaTeX:** drop cell $m$ if $m < m_{\mathrm{floor}}(I)$.

| Instrument | $m_{\mathrm{floor}}$ | Comment in CLI / tests |
|------------|---------------------:|------------------------|
| violin | 79 | G5 |
| viola | 72 | C5 |
| cello | 60 | C4 |
| double bass | 52 | E3 |

**E. Symbols:** concert MIDI. Intent (source comments): first **measured Orchidea chromatic harmonic**, not the catalogue `harmonics_sounding_low` (which is lower: violin 67, viola 60, cello 48, bass 40).

**F. Layman:** Harmonic books throw away notes below a per-instrument cutoff so low “harmonics” do not enter $L$ or Media.

**G. Specialist:** Discrete policy, not a nodal-frequency model. Two different floors exist (catalogue vs pipeline). Cells below the floor are deleted, not set to missing-with-flag.

**H. Downstream:** `transfer_or_copy(..., min_midi=...)`; preparatory `_anchors(..., min_midi=floor)`.

**I. Tests:** `test_harm_floor_by_instrument`; `test_harmonics_floor_drops_notes_below_c5`.

---

### M-007 — Measured-ceiling MIDI (production)

**A. Status:** production for reporting status and Zenodo shading.

**B. Source:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py), `measured_ceiling_midi`, lines 187–197. Default constant `DEFAULT_MEASURED_CEILING_MIDI = 100` in [ste_lab/catalog.py](../ste_lab/catalog.py) line 48.

**C. Snippet:**

```python
def measured_ceiling_midi(layers: dict[tuple[str, str], Layer], instrument: str = "") -> int:
    measured = [
        m
        for lg in layers.values()
        for m, c in lg.cells.items()
        if (c.origin or "").lower() == "measured"
    ]
    if measured:
        return max(measured)
    spec = resolve_instrument(instrument) if instrument else None
    return spec.sounding_high if spec else DEFAULT_MEASURED_CEILING_MIDI
```

**D. LaTeX:**

$$
C =
\begin{cases}
\max\{m:\ \mathrm{origin}(m)=\texttt{measured}\} & \text{if that set is nonempty}\\
h_{\mathrm{sounding}}(I) & \text{else if catalogue exists}\\
100 & \text{otherwise}
\end{cases}
$$

**E. Symbols:** $C$ ceiling MIDI. “Measured” is the **origin tag**, not an independent measurement oracle.

**F. Layman:** The highest note marked measured becomes the ceiling. Notes above it are treated as register-extrapolated for reporting.

**G. Specialist:** If modelled cells are mis-tagged `measured`, the ceiling rises and tails enter “measured range.” Default 100 is a violin-era constant (G7-ish), not a woodwind fact.

**H. Downstream:** M-008; Summary filters; AcousticTable shading.

**I. Tests:** `test_bass_clarinet_grid_and_ceiling`; `test_pipeline_regression` bass-clarinet ceiling assertion.

---

### M-008 — Reporting status vs ceiling (production)

**A. Status:** production.

**B. Source:** [ste_lab/session.py](../ste_lab/session.py) `Layer.place` lines 93–97; [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) `_status` lines 145–146; prediction path in [run_ste_from_preparatory.py](../run_ste_from_preparatory.py) lines 459–462 (`midi > 88`).

**C. Snippet:**

```python
        if not reporting_status:
            if origin == "modelled_on_extrapolated_anchor" or pitch.midi > self.measured_ceiling_midi:
                reporting_status = "above_measured_ceiling"
            else:
                reporting_status = "measured_range"
```

```python
def _status(midi: int, ceiling: int) -> str:
    return "above_measured_ceiling" if midi > ceiling else "measured_range"
```

**D. LaTeX:**

$$
\mathrm{status}(m)=
\begin{cases}
\texttt{above\_measured\_ceiling} & m>C\ \text{or origin is modelled-on-extrapolated-anchor}\\
\texttt{measured\_range} & \text{otherwise}
\end{cases}
$$

Prediction-only IOWA cells additionally force $m>88$ to `above_measured_ceiling`.

**E. Symbols:** $C$ from M-007 or layer default 100. Units: MIDI.

**F. Layman:** A flag saying “this pitch is above the last trusted measured note.”

**G. Specialist:** Status is **not** a statistical interval. Summary sheets claim to exclude these cells; Media construction (M-027) does **not** automatically drop them unless `reporting_status` is later filtered.

**H. Downstream:** QA `above_measured_ceiling`; Summary_Measured_Range filters; Zenodo fill colour.

**I. Tests:** Acoustic-table tests check status columns; no isolated formula test.

---

### M-009 — Chromatic note grid for woodwind/Zenodo tables (production)

**A. Status:** production.

**B. Source:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py), `_note_grid`, lines 162–184.

**C. Snippet:**

```python
    if technique == "harmonics" and present:
        return list(range(min(present), max(present) + 1))
    lo = spec.sounding_low if spec else (min(present) if present else 55)
    hi = spec.sounding_high if spec else (max(present) if present else 107)
    extra = measured or present
    if extra:
        lo = min(lo, min(extra))
        hi = max(hi, max(extra))
    return list(range(lo, hi + 1))
```

**D. LaTeX:** inclusive integer range. Harmonics: $[\min m_{\mathrm{present}},\max m_{\mathrm{present}}]$. Else expand catalogue $[h_{\mathrm{lo}},h_{\mathrm{hi}}]$ (fallback $55$–$107$) to include measured or present extrema.

**E. Symbols:** MIDI integers. Fallback $55$–$107$ is the old violin box (comment).

**F. Layman:** Decide which chromatic notes appear as rows.

**G. Specialist:** Iowa notes beyond catalogue `sounding_high` **are** included (intentional; tests require this). Empty layers fall back to 55–107, which is not a bass-clarinet range.

**H. Downstream:** AcousticTable and calibration grids.

**I. Tests:** `test_bass_clarinet_grid_and_ceiling`; `test_c_notes_beyond_catalog_sounding_high`.

---

### Stage B — Ingested metric, merge, and parsing

### M-010 — F-061 / spectral_mass ingest (production; value not computed here)

**A. Status:** production ingest. **No project formula for the metric itself.**

**B. Source:** [run_ste_effects_batch.py](../run_ste_effects_batch.py), `_extract_mass` / `load_spectral_mass`, lines 126–159. Violin-4 variant in [run_ste_effects_batch_violin4.py](../run_ste_effects_batch_violin4.py) uses a normalised column key `spectral_mass`.

**C. Snippet:**

```python
        pitch = parse_pitch(row.get(note_col))
        try:
            value = float(row.get(mass_col))
        except (TypeError, ValueError):
            continue
        if pitch is None or value <= 0:
            continue
        out[pitch.midi] = (pitch.label, value)
```

**D. LaTeX:** no generating equation. Implemented rule: keep $(m,y)$ iff $y$ parses as float and $y>0$.

**E. Symbols:** $y$ is whatever the compiled workbook stored under `spectral_mass`. Comments call it F-061 Combined Density Metric. **Units are not defined in this repository.** They are **not** converted to dB.

**F. Layman:** The program copies a positive density number next to a note name from an Excel sheet.

**G. Specialist:** Provenance of $y$ is an **external compile**. This audit cannot reconstruct F-061. Treating $y$ as SPL or “brightness” is unsupported here (`gui_app.py` help text says F-061 is not brightness).

**H. Downstream:** all ratios $L=\ln(y_t/y_s)$ inherit this unitless-or-unknown scale. Logs are natural logs of the **same** stored quantity.

**I. Tests:** discovery/regression tests check that values appear; they do not recompute F-061.

---

### M-011 — Positive finite cell filter (production)

**A. Status:** production, many call sites. Canonical helpers: `_finite` in [run_ste_from_preparatory.py](../run_ste_from_preparatory.py) lines 141–146; `apply_log_ratio` / `finite_pairs` in [ste_lab/calibration.py](../ste_lab/calibration.py).

**C. Snippet:**

```python
def _finite(v) -> bool:
    try:
        x = float(v)
    except (TypeError, ValueError):
        return False
    return math.isfinite(x) and x > 0
```

**D. LaTeX:** keep $x$ iff $x\in\mathbb{R}$, $x>0$, and $x$ is finite.

**E. Symbols:** same stored metric as M-010.

**F. Layman:** Zero, negative, and blank cells are ignored.

**G. Specialist:** Combined density is assumed strictly positive. Zeros never enter logs. There is no missing-data model (no imputation at ingest).

**H. Downstream:** every ratio and mean.

**I. Tests:** `test_empty_and_nan` (agreement metrics).

---

### M-012 — Empirical merge: union, conflict drop, no average (production)

**A. Status:** production in preparatory rebuild.

**B. Source:** [ste_lab/empirical.py](../ste_lab/empirical.py), `merge_measured_specs`, lines 119–137.

**C. Snippet:**

```python
            values = [item.value for item in group]
            if all(values_equivalent(values[0], v) for v in values[1:]):
                curve[midi] = (group[0].label or midi_to_label(midi), values[0])
                report.n_exported_empirical_midis += 1
                continue
            report.n_conflicts += 1
            report.n_dropped += 1
            reason = {
                "midi": midi,
                "reason": "EMPIRICAL_CONFLICT",
                "values": values,
                "sources": [str(item.source_path) for item in group],
            }
```

**D. LaTeX:** for each $(c,t,d,m)$ group $Y=\{y_j\}$:

$$
\hat y(m)=
\begin{cases}
y_1 & \text{if } |y_j-y_1|\le \tau(y_1,y_j)\ \forall j\\
\text{dropped} & \text{otherwise}
\end{cases}
$$

No $\bar y$, no winner-by-filesize.

**E. Symbols:** $\tau$ from M-013. $c,t,d$ collection, technique, dynamic.

**F. Layman:** If two files disagree on the same note, that note is thrown away, not averaged.

**G. Specialist:** Conservative union. Dependence across files is not modelled. “Longest-file-wins is prohibited” (module docstring) is implemented.

**H. Downstream:** preparatory measured curves.

**I. Tests:** `test_union_not_longest_wins`; `test_conflict_not_averaged`.

---

### M-013 — Relative equivalence test (production)

**A. Status:** production (merge) and reusable helper.

**B. Source:** [ste_lab/empirical.py](../ste_lab/empirical.py) lines 12–13, 49–50.

**C. Snippet:**

```python
REL_TOL = 1e-6
ABS_TOL = 1e-9
```

```python
    return abs(a - b) <= max(abs_tol, rel * max(abs(a), abs(b), 1.0))
```

**D. LaTeX:**

$$|a-b|\le \max\bigl(\varepsilon_{\mathrm{abs}},\,\rho\cdot\max(|a|,|b|,1)\bigr)$$

with $\rho=10^{-6}$, $\varepsilon_{\mathrm{abs}}=10^{-9}$.

**E. Symbols:** same units as $y$. The floor $1$ inside $\max$ means relative tolerance is at least $\rho$ in absolute terms when both $|a|,|b|<1$.

**F. Layman:** Tiny numerical differences count as “the same number.”

**G. Specialist:** Not a statistical test. The $\max(\cdot,1)$ floor is a numerical heuristic that loosens relative comparison for sub-unit densities.

**H. Downstream:** M-012.

**I. Tests:** conflict tests use clearly distinct values.

---

### M-014 — Locale-aware decimal parse (production)

**A. Status:** production GUI/paste.

**B. Source:** [ste_lab/paste.py](../ste_lab/paste.py), `parse_number`, lines 71–85. Anchor floats in [ste_lab/transfer.py](../ste_lab/transfer.py) `parse_anchor_block` lines 352–353 also replace `,` with `.`.

**C. Snippet:**

```python
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    elif s.count(".") > 1 and s.count(",") == 1:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None
```

**D. LaTeX:** parsing map, not an estimator. `12,3` $\to$ $12.3$; `1.234,5` $\to$ $1234.5$.

**E. Symbols:** none.

**F. Layman:** European commas become decimal points.

**G. Specialist:** Ambiguous `1,234` (thousands) becomes $1.234$. Not IEEE parsing.

**H. Downstream:** pasted CDM and PI95 columns.

**I. Tests:** No direct test identified.

---

### Stage C — Log-ratio $L$ and technique transfer

### M-015 — Safe natural log-ratio (production)

**A. Status:** production kernel.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py), `_safe_log_ratio`, lines 37–40. Preparatory duplicate: [build_ste_preparatory.py](../build_ste_preparatory.py) line 466 `math.log(tgt / src)`.

**C. Snippet:**

```python
def _safe_log_ratio(target: float, source: float) -> Optional[float]:
    if source <= 0 or target <= 0:
        return None
    return float(np.log(target / source))
```

**D. LaTeX:**

$$
L=\ln\frac{y_{\mathrm{tgt}}}{y_{\mathrm{src}}}
\quad\text{if }y_{\mathrm{tgt}}>0\text{ and }y_{\mathrm{src}}>0;\quad
\text{else undefined.}
$$

**E. Symbols:** $y$ ingested metric (M-010). $\ln=\log_e$ (`numpy.log`). $L$ is **dimensionless**. It is **not** a decibel ($20\log_{10}$ or $10\log_{10}$ are absent).

**F. Layman:** $L$ says how many natural-log steps the effect (or family target) is above or below the source curve at that note. $L=0$ means same value; $L=\ln 2$ means twice as large.

**G. Specialist:** Multiplicative model $y_{\mathrm{tgt}}=y_{\mathrm{src}}e^{L}$. No error model. Undefined on non-positive $y$ (no continuity correction). Comments elsewhere write $L=\ln(\mathrm{target}/\mathrm{source})$ consistently with this implementation.

**H. Downstream:** M-016–M-021, preparatory `Anchors_all` column `L`.

**I. Tests:** `test_anchor_identity_source_times_exp_L`.

---

### M-016 — Register-dependent $L$ with edge hold (production)

**A. Status:** production. Shared by technique transfer and `two_log_ratio`.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py), `_interp_log_ratio`, lines 43–76.

**C. Snippet:**

```python
    usable = [a for a in anchors if _safe_log_ratio(a.target_value, a.source_value) is not None]
    if not usable:
        return {}
    xs = np.array([a.midi for a in usable], dtype=float)
    ys = np.array([_safe_log_ratio(a.target_value, a.source_value) for a in usable], dtype=float)
    order = np.argsort(xs)
    xs, ys = xs[order], ys[order]
    if len(xs) == 1:
        return {m: float(ys[0]) for m in midis}
    lo, hi = float(xs[0]), float(xs[-1])
    y_lo, y_hi = float(ys[0]), float(ys[-1])
    if len(xs) == 2:
        interior = {m: float(np.interp(m, xs, ys)) for m in midis if lo <= m <= hi}
    else:
        fn = PchipInterpolator(xs, ys, extrapolate=False)
        interior = {m: float(fn(m)) for m in midis if lo <= m <= hi}
    out: dict[int, float] = {}
    for m in midis:
        if m < lo:
            out[m] = y_lo
        elif m > hi:
            out[m] = y_hi
        else:
            out[m] = interior[m]
    return out
```

**D. LaTeX:** let anchors be pairs $(m_i,L_i)$ sorted by $m$. $n=|\{i\}|$.

$$
L(m)=
\begin{cases}
L_1 & n=1\\
\mathrm{lerp}(m;m_1,m_2,L_1,L_2) & n=2,\ m\in[m_1,m_2]\\
\mathrm{PCHIP}(m;\{m_i,L_i\}) & n\ge 3,\ m\in[m_{\min},m_{\max}]\\
L(m_{\min}) & m<m_{\min}\\
L(m_{\max}) & m>m_{\max}
\end{cases}
$$

PCHIP internals: **L-001**. Two-point lerp: **L-002**.

**E. Symbols:** $m$ MIDI (semitone index). $L(m)$ dimensionless. Duplicate MIDI after sort is not de-duplicated; SciPy PCHIP forbids duplicate $x$ (would error).

**F. Layman:** Between measured pairs, $L$ is smoothly interpolated. Outside, the last measured $L$ is copied (held), not allowed to dive.

**G. Specialist:** Edge hold is an explicit **anti-extrapolation** policy. Comments (lines 46–50) state that PCHIP `extrapolate=True` previously produced $L\approx -20$ to $-40$ and near-zero high notes after clipping. This is a heuristic, not a physical register law. Interior PCHIP is shape-preserving cubic; it is not a GP.

**H. Downstream:** M-018, M-021. GUI fill (M-032) is a **different** interpolator (values, not $L$; PCHIP may extrapolate).

**I. Tests:** `test_edge_L_is_held_not_cubically_extrapolated`.

---

### M-017 — Multiplicative apply with $\pm 3$ clip (production)

**A. Status:** production.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py) lines 111–114; [ste_lab/calibration.py](../ste_lab/calibration.py) `apply_log_ratio` / `L_CLIP`, lines 37, 203–214.

**C. Snippet:**

```python
            ratio = float(np.exp(np.clip(L[midi], -3.0, 3.0)))
            value = cell.value * ratio
            if not np.isfinite(value) or value <= 0:
                continue
```

```python
L_CLIP = 3.0
```

```python
        value = src * math.exp(min(L_CLIP, max(-L_CLIP, L[midi])))
```

**D. LaTeX:**

$$\tilde L=\mathrm{clip}(L,-3,3),\qquad \hat y = y_{\mathrm{src}}\,e^{\tilde L}$$

keep $\hat y$ iff $\hat y>0$ and finite.

$$\mathrm{clip}(L,-3,3)=\min\bigl(3,\max(-3,L)\bigr)$$

implies $e^{\tilde L}\in[e^{-3},e^{3}]\approx[0.049787,20.085537]$.

**E. Symbols:** $y_{\mathrm{src}}$ source-layer value; $\hat y$ modelled target. Same unknown units as M-010.

**F. Layman:** Multiply the source curve by a ratio. The ratio is not allowed to be smaller than about $1/20$ or larger than about $20$.

**G. Specialist:** Hard safeguard, not a statistical residual bound. Extreme measured pairs are distorted toward $\pm 3$. Inverse is $\hat y/y_{\mathrm{src}}=e^{\tilde L}$, not the unclipped pair ratio when $|L|>3$.

**H. Downstream:** every transferred cell.

**I. Tests:** `test_extreme_ratio_is_clipped_at_exp_3`; `test_anchor_identity_source_times_exp_L`.

---

### M-018 — Technique transfer `log_ratio` (production)

**A. Status:** production default for string effects and GUI technique transfer.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py), `technique_transfer`, lines 106–133. Wrapper [run_ste_effects_batch.py](../run_ste_effects_batch.py) `transfer_or_copy` lines 426–450 retags non-anchor origins.

**C. Snippet:**

```python
    if method == "log_ratio":
        L = _interp_log_ratio(anchors, list(source.cells))
        for midi, cell in source.cells.items():
            if midi not in L:
                continue
            ratio = float(np.exp(np.clip(L[midi], -3.0, 3.0)))
            value = cell.value * ratio
            if not np.isfinite(value) or value <= 0:
                continue
            origin = "technique_transfer"
            comment = f"from {source.technique}"
            anchor_midis = [a.midi for a in anchors]
            for a in anchors:
                if a.midi == midi and copy_anchors:
                    origin = "measured"
                    value = a.target_value
                    result.n_copied += 1
                    break
            else:
                result.n_transferred += 1
```

**D. LaTeX:**

$$
\hat y(m)=
\begin{cases}
y^{\mathrm{tgt}}_{\mathrm{anchor}}(m) & \text{if }m\text{ is an anchor and }\texttt{copy\_anchors}\\
y_{\mathrm{src}}(m)\,e^{\mathrm{clip}(L(m),-3,3)} & \text{otherwise, if defined}
\end{cases}
$$

Batch `transfer_or_copy` uses `copy_anchors=False` and overwrites origin to `modelled_IOWA_anchored` / `modelled_ORCHIDEA_anchored`. Recorded Orchidea effect dynamics **bypass** this function (`orchidea_recorded_layer`: L is not applied).

**E. Symbols:** $L(m)$ from M-016 on effect/ordinario pairs. $y_{\mathrm{src}}$ is typically IOWA or ORCH **ordinario/arco**.

**F. Layman:** Teach the computer the effect-to-ordinario ratio at a few notes, then multiply the whole ordinary curve by that ratio. Where Orchidea actually recorded the effect, those recordings replace the model.

**G. Specialist:** Shared-$L$ construction: the **same** $L$ is applied to IOWA and ORCH spines. IOWA effect curves are **not** an independent experiment (`SHARED_L_MEDIA_CAVEAT`). `copy_anchors=True` labels copied targets `measured` even though they came from the anchor table (may or may not be a raw recording).

**H. Conditions:** batch requires $\ge 3$ anchors unless a recorded Orchidea curve exists. GUI warns if $n_{\mathrm{anchors}}<3$ (`technique_without_anchors`) but can proceed after confirm.

**I. Tests:** `test_exported_books_obey_L_and_evidence_rules`; `test_recorded_orch_dynamic_is_not_replaced_by_L`; `test_iowa_same_dynamic_is_modelled_not_inherited`.

---

### M-019 — Technique transfer `linear_ratio` (optional / GUI)

**A. Status:** optional. GUI combo includes it; batch always passes `"log_ratio"`.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py) lines 135–149.

**C. Snippet:**

```python
    if method == "linear_ratio":
        ratios = [a.target_value / a.source_value for a in anchors if a.source_value > 0]
        scale = float(np.median(ratios)) if ratios else 1.0
        for midi, cell in source.cells.items():
            hit = next((a for a in anchors if a.midi == midi), None)
            if hit:
                target.cells[midi] = Cell(midi, cell.note_label, hit.target_value, "measured", "anchor")
                result.n_copied += 1
            else:
                target.cells[midi] = Cell(
                    midi, cell.note_label, cell.value * scale, "technique_transfer", f"global ratio {scale:.4f}"
                )
```

**D. LaTeX:** $r_i=y^{\mathrm{tgt}}_i/y^{\mathrm{src}}_i$ for $y^{\mathrm{src}}_i>0$. $s=\mathrm{median}\{r_i\}$ or $1$ if empty. $\hat y(m)=y^{\mathrm{tgt}}_{\mathrm{anchor}}(m)$ on anchors; else $s\cdot y_{\mathrm{src}}(m)$. Median: L-style NumPy median (linear interpolation at even $n$ — NumPy default).

**E. Symbols:** global scalar $s$, not register-dependent.

**F. Layman:** One typical ratio for the whole instrument, not a curve of ratios.

**G. Specialist:** Ignores register structure that M-016 exists to capture. Not the production batch path.

**H. Downstream:** GUI only unless a caller passes `method="linear_ratio"`.

**I. Tests:** No direct test identified.

---

### M-020 — Preparatory exported $L$ columns (offline generator)

**A. Status:** offline / preparatory workbook generator.

**B. Source:** [build_ste_preparatory.py](../build_ste_preparatory.py) `_anchors` lines 448–468; grid $L_*$ lines 956–974.

**C. Snippet:**

```python
                "L": math.log(tgt / src),
```

```python
            if _finite_cell(rec, spine_key) and _finite_cell(rec, col):
                rec[l_key] = math.log(rec[col] / rec[spine_key])
            else:
                rec[l_key] = float("nan")
```

```python
                rec["L_sordino_McGill"] = math.log(rec["McGill_sordino_mf"] / rec[f"McGill_{spine}_mf"])
```

**D. LaTeX:** same as M-015 on each overlapping MIDI. `Media_{spine}_mf=(y_{\mathrm{IOWA}}+y_{\mathrm{ORCH}})/2` when both finite (line 948) — arithmetic, see M-028.

**E. Symbols:** `sourceCDM` = teacher ordinario; `targetCDM` = effect. McGill sordino $L$ is contextual, not Media.

**F. Layman:** The preparatory Excel writes the log-ratio so a human can see it; the runner later rebuilds $L$ from `sourceCDM`/`targetCDM`, not by reading the `L` column.

**G. Specialist:** `load_preparatory` builds `Anchor(midi, sourceCDM, targetCDM)` and ignores the stored `L` column. Recomputing $L$ from floats should match except for display rounding in `STE_Lab_paste`.

**H. Downstream:** `Anchors_all` consumed by `run_ste_from_preparatory`.

**I. Tests:** preparatory sheet tests; L identity tests use reconstructed anchors.

---

### Stage D — Family transfer (woodwinds) and GUI family map

### M-021 — Two-log-ratio family transfer (production default)

**A. Status:** production for woodwinds when `calibration.yaml` has `transfer.method: two_log_ratio` (default).

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `two_log_ratio_transfer`, lines 295–325; applied in [build_ste_preparatory.py](../build_ste_preparatory.py) `write_woodwind_media` lines 572–585.

**C. Snippet:**

```python
    """ORCH_target(m) = ORCH_clarinet(m) × exp(L_instr(m)).

    L_coll(m) = ln(ORCH_clarinet / IOWA_clarinet)  (same instrument)
    L_instr(m) = ln(IOWA_target / IOWA_clarinet)   (same collection)

    L_instr uses the target Iowa curve, so IOWA is excluded from validation.
    """
    coll_anchors = _anchors(iowa_clarinet, orch_clarinet)
    instr_anchors = _anchors(iowa_clarinet, iowa_target)
    L_coll = _interp_log_ratio(coll_anchors, midis) if coll_anchors else {}
    L_instr = _interp_log_ratio(instr_anchors, midis) if instr_anchors else {}
    curve = apply_log_ratio(orch_clarinet, L_instr, midis) if L_instr else {}
```

**D. LaTeX:** at overlapping MIDI used as anchors,

$$
L_{\mathrm{coll}}(m)=\ln\frac{y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)}{y^{\mathrm{IOWA}}_{\mathrm{sib}}(m)},\qquad
L_{\mathrm{instr}}(m)=\ln\frac{y^{\mathrm{IOWA}}_{\mathrm{tgt}}(m)}{y^{\mathrm{IOWA}}_{\mathrm{sib}}(m)}
$$

then interpolate/hold as M-016. **Executed** transfer:

$$\hat y^{\mathrm{ORCH}}_{\mathrm{tgt}}(m)=y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)\,e^{\mathrm{clip}(L_{\mathrm{instr}}(m),-3,3)}$$

$L_{\mathrm{coll}}$ is stored on `TransferEstimate` and written to `Transfer_fit` / calibration sheets. It is **not** multiplied into $\hat y$. Do **not** add an extra factor $e^{L_{\mathrm{coll}}}$.

At a **matched unclipped anchor** the unused diagnostic identity

$$
y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)\,e^{L_{\mathrm{instr}}(m)}
=
y^{\mathrm{IOWA}}_{\mathrm{tgt}}(m)\,e^{L_{\mathrm{coll}}(m)}
=
\frac{y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)\,y^{\mathrm{IOWA}}_{\mathrm{tgt}}(m)}{y^{\mathrm{IOWA}}_{\mathrm{sib}}(m)}
$$

holds because the predictor already starts from the ORCH donor. This identity is **not** claimed for interpolated or clipped MIDI.

If `instr_anchors` is empty: $L_{\mathrm{instr}}$ status `UNDETERMINED`, $\mathrm{curve}=\{\}$.

**E. Symbols:** “clarinet” in names is the **family donor** (`FAMILY_TRANSFER_DONOR`, e.g. bass clarinet $\to$ clarinet). $y$ as M-010. Iowa target used to fit $L_{\mathrm{instr}}$ is listed in `validation_excluded=["IOWA"]`.

**F. Layman:** Guess the missing instrument’s Orchidea curve by taking the cousin’s Orchidea curve and stretching it by the Iowa family ratio (bass vs clarinet on Iowa).

**G. Specialist:** Identifiability: $L_{\mathrm{instr}}$ is fitted on Iowa, then applied to Orchidea cousin. That is a **transport** assumption (instrument ratio portable across collections), not a fitted collection×instrument interaction. $L_{\mathrm{coll}}$ is a **diagnostic collection ratio**, not a second production multiplier. Iowa cannot validate this transfer (leakage). No standard errors.

**H. Downstream:** woodwind ORCH layer / `orchidea_family_transfer_estimate`; `build_final_rows` transfer column; Media when Iowa missing at a MIDI.

**I. Tests:** `TestTwoLogRatioNoLeakage`; `test_iowa_excluded_from_validation_list`; `test_woodwind_orch_is_transferred_not_measured`.

---

### M-022 — Even-length average median of ratios (production helper)

**A. Status:** production helper. Used by legacy scale, collection offset, and shape offset (`_median`).

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py) `legacy_target_scale` lines 178–191; `_median` 378–385; `collection_offset` 438–447.

**C. Snippet:**

```python
    ratios.sort()
    mid = len(ratios) // 2
    if len(ratios) % 2:
        return float(ratios[mid])
    return float(0.5 * (ratios[mid - 1] + ratios[mid]))
```

**D. LaTeX:** for sorted $r_{(1)}\le\cdots\le r_{(n)}$,

$$
\mathrm{median}=
\begin{cases}
r_{((n+1)/2)} & n\text{ odd}\\
\frac{1}{2}\bigl(r_{(n/2)}+r_{(n/2+1)}\bigr) & n\text{ even}
\end{cases}
$$

Empty $\to 1.0$ in `legacy_target_scale`; empty $\to$ `None` in `_median` / `collection_offset`.

**E. Symbols:** $r$ are **ratios** $y_a/y_b$ or log-residuals, depending on caller.

**F. Layman:** Typical middle ratio, averaging the two middle numbers if there is an even count.

**G. Specialist:** Standard sample median. **Not** the same as `agreement_metrics` MedAE, which uses `sorted(abs_err)[n // 2]` only (M-035).

**H. Downstream:** M-023, M-036, M-037.

**I. Tests:** legacy combined-estimate test; shape-metric analytical fixture.

---

### M-023 — Legacy scalar family transfer (optional / deprecated)

**A. Status:** optional. `transfer.method = legacy_target_calibrated`. Outputs labelled `COMBINED_ESTIMATE`, never `TRANSFERRED`.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py) `legacy_scalar_transfer` lines 278–301.

**C. Snippet:**

```python
    scale = legacy_target_scale(orch_clarinet, iowa_target)
    curve = {
        int(m): float(orch_clarinet[m] * scale)
        for m in midis
        if lo <= m <= hi and orch_clarinet.get(m, 0) > 0
    }
```

with

```python
        target[m] / val
        for m, val in donor.items()
        if val > 0 and target.get(m, 0) > 0 ...
```

**D. LaTeX:**

$$s=\mathrm{median}_m\frac{y^{\mathrm{IOWA}}_{\mathrm{tgt}}(m)}{y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)},\qquad
\hat y(m)=s\,y^{\mathrm{ORCH}}_{\mathrm{sib}}(m)\ \text{for }m\in[\ell,h]\text{ and }y_{\mathrm{sib}}(m)>0.$$

**E. Symbols:** $s$ dimensionless. $[\ell,h]$ are catalogue `sounding_low`/`sounding_high` passed by the caller (unlike default two-ratio, which uses `sorted(donor)` and does not clip to catalogue high).

**F. Layman:** One median stretch factor so the cousin Orchidea curve sits on the Iowa target.

**G. Specialist:** Calibrates donor **on the target** (leakage). Module docstring: numerical family-transfer used through 1.2.8. Not a log-register model.

**H. Downstream:** only if YAML method is switched.

**I. Tests:** `test_legacy_scalar_is_combined_estimate`.

---

### M-024 — Philharmonia/McGill pairing $L_{\mathrm{instr}}$ when host Iowa is absent (optional production branch)

**A. Status:** production fallback inside `write_woodwind_media` when **no** measured Iowa of the host exists, but sibling Iowa/Orchidea do.

**B. Source:** [build_ste_preparatory.py](../build_ste_preparatory.py) `_pairing_L_instr` lines 515–530; apply lines 595–618.

**C. Snippet:**

```python
    """Same-collection L_instr = ln(target / sibling). Philharmonia first, then McGill."""
    ...
    for coll in ("PHIL", "MCGILL"):
        tgt = _as_map(_get(measured, coll, "ordinario", dyn)) or _as_map(
            _get(measured, coll, "ordinario", "mf")
        )
        sib = _as_map(_get(donor_measured, coll, "ordinario", dyn)) or _as_map(
            _get(donor_measured, coll, "ordinario", "mf")
        )
        anchors = _anchors(sib, tgt)
        if anchors:
            return _interp_log_ratio(anchors, midis), anchors, coll
```

```python
            orch[dyn] = apply_log_ratio(donor, L_instr, midis) if donor else {}
            iowa[dyn] = apply_log_ratio(sibling, L_instr, midis) if sibling else {}
```

**D. LaTeX:** first collection $c\in\{\mathrm{PHIL},\mathrm{MCGILL}\}$ with overlap:

$$L_{\mathrm{instr}}(m)=\ln\frac{y^{c}_{\mathrm{tgt}}(m)}{y^{c}_{\mathrm{sib}}(m)}\ \text{(interpolated as M-016)}$$

$$\hat y^{\mathrm{ORCH}}_{\mathrm{tgt}}=y^{\mathrm{ORCH}}_{\mathrm{sib}}e^{\mathrm{clip}(L_{\mathrm{instr}},\pm3)},\qquad
\hat y^{\mathrm{IOWA}}_{\mathrm{tgt}}=y^{\mathrm{IOWA}}_{\mathrm{sib}}e^{\mathrm{clip}(L_{\mathrm{instr}},\pm3)}$$

Iowa origin becomes `family_transfer` (not `measured`). Dynamic fallback: if $c$ lacks this dynamic, use that collection’s **mf**.

**E. Symbols:** PHIL/McGill $y$ are still ingested F-061, not a different physical quantity — but they are a different **codec** (MP3 vs WAV; see `COLLECTION_CODEC`).

**F. Layman:** If this instrument was never measured at Iowa, copy the family ratio from Philharmonia (else McGill) onto both the cousin Iowa and cousin Orchidea curves.

**G. Specialist:** Both “IOWA” and “ORCH” columns can be **purely modelled**. They share one $L$ and are dependent. Validation excludes the pairing collection. Codec/collection shift is unmodelled.

**H. Downstream:** generated media workbook; tests `test_phil_ratio_writes_iowa_and_orch_without_host_iowa`.

**I. Tests:** that regression test; `test_woodwind_orch_fallback_without_origins`.

---

### M-025 — GUI family concert-pitch / register-relative map (optional)

**A. Status:** optional GUI. Not used by woodwind `two_log_ratio` batch.

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py), `family_transfer`, lines 154–226.

**C. Snippet:**

```python
    ratios = [a.target_value / a.source_value for a in anchors if a.source_value > 0]
    if ratios:
        scale = float(np.median(ratios))
```

```python
                target.cells[midi] = Cell(
                    midi, midi_to_label(midi), cell.value * scale, "family_transfer", f"concert × {scale:.4f}"
                )
```

```python
            offset = midi - src_low
            tgt_midi = tgt_low + offset
```

**D. LaTeX:** $s=\mathrm{median}(y_{\mathrm{tgt}}/y_{\mathrm{src}})$ or $1$.

- `concert_pitch`: copy/scale cells with the **same** sounding MIDI inside the target catalogue range; anchors overwrite as `measured`.
- `register_relative`: $m'=\ell_{\mathrm{tgt}}+(m-\ell_{\mathrm{src}})$, then $\hat y(m')=s\,y(m)$.

Harmonic target range uses M-005.

**E. Symbols:** $\ell$ sounding_low. $s$ dimensionless.

**F. Layman:** Either line the instruments up at the same concert notes, or slide one range so their bottom notes match, then scale.

**G. Specialist:** Register-relative assumes equal-tempered index alignment is musically meaningful. Cross-family use is QA-blocked as critical but not hard-stopped without the GUI confirm path.

**H. Downstream:** GUI only.

**I. Tests:** No direct numerical test identified.

---

### Stage E — Combination, Media, and cell overwrite

### M-026 — Combine empirical and modelled (production)

**A. Status:** production. Default YAML `empirical_only`.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `combine_empirical_and_modelled`, lines 362–397; `validate_combination_config` lines 150–168. Supported YAML names: `empirical_only`, `equal_weight`, `legacy_equal_weight`. `fixed_weight` and `validation_optimised` now **raise** `ConfigurationError`. `optimal_weight()` remains a research helper and is not a production workflow.

**C. Snippet:**

```python
    validate_combination_config(cfg)
    method = cfg.combination_method
    w_e = cfg.empirical_weight if empirical_weight is None else float(empirical_weight)
    if not math.isfinite(w_e) or w_e < 0.0 or w_e > 1.0:
        raise ConfigurationError(f"empirical_weight={w_e!r} must be finite and in [0, 1].")
    w_t = 1.0 - w_e
    has_e = empirical is not None and empirical > 0 and math.isfinite(empirical)
    has_t = transfer is not None and transfer > 0 and math.isfinite(transfer)
    if method == "legacy_equal_weight" or method == "equal_weight":
        if method == "equal_weight":
            w_e, w_t = 0.5, 0.5
        if has_e and has_t:
            return w_e * empirical + w_t * transfer, method, w_e, w_t
        if has_e:
            return empirical, method, 1.0, 0.0
        if has_t:
            return transfer, method, 0.0, 1.0
        return None, method, w_e, w_t
    if has_e:
        return empirical, "empirical_only", 1.0, 0.0
    if has_t:
        return transfer, "transfer_only", 0.0, 1.0
    return None, method, w_e, w_t
```

**D. LaTeX:**

$$
(\hat y, w_E, w_T)=
\begin{cases}
(w_E y_E + w_T y_T,\ w_E,\ w_T) & \text{equal-weight family and both present}\\
(y_E,\ 1,\ 0) & y_E\text{ present and method not forcing a mix, or only }y_E\\
(y_T,\ 0,\ 1) & \text{only }y_T\\
(\emptyset,\ \cdot,\ \cdot) & \text{neither}
\end{cases}
$$

where $w_T=1-w_E$ is the **authoritative** legacy rule. `cfg.transfer_weight` is accepted only when it equals $1-w_E$; a contradictory value raises `ConfigurationError`. $w_E=w_T=1/2$ for `equal_weight`. Per-row exported weights are the **effective** $(w_E,w_T)$; `Configured_*` columns store the YAML defaults.

Unsupported or misspelled method names raise `ConfigurationError`. They no longer silently become `empirical_only`.

A blended output (`equal_weight` / `legacy_equal_weight` with both inputs and $0<w_E<1$) is labelled `COMBINED_ESTIMATE`. The raw empirical input may remain `MEASURED`. Existing combined-estimate QA: relative disagreement $>0.5$ $\to$ `REVIEW_REQUIRED`, else `MODERATE`. No new confidence threshold was invented.

**E. Symbols:** $y_E$ Iowa (or tagged empirical) value; $y_T$ transfer estimate. Same metric units.

**F. Layman:** By default, if Iowa measured the note, that number is the answer. The model is used only where Iowa is missing.

**G. Specialist:** `empirical_only` is a hard override, not a Bayes average. Methodology text that “weighting applies only where no target measurement exists” **contradicts** `equal_weight` / `legacy_equal_weight`, which **do** replace a measurement by a convex combination. `optimal_weight` (M-038) is not called from this function.

**H. Downstream:** `build_final_rows` `final_value`; woodwind Media via `production_media_of` (M-027).

**I. Tests:** `TestPrecedenceAndWeights`; `test_export_consistency.py`.

---

### M-027 — Production Media resolver (strings vs woodwind)

**A. Status:** production. Shared resolver: [ste_lab/media_policy.py](../ste_lab/media_policy.py). Generic `media_of` remains an explicit alternative averager, not the production policy.

**B. Source:** `resolve_iowa_orch_pair` / `production_media_of` lines 62–196. Callers: `run_ste_from_preparatory.py`, GUI `production_average_for_layer`, Zenodo `_resolved_media`.

**C. Snippet:**

```python
    if policy == "string_mean":
        vals = [v for v in (iowa, orch) if v is not None]
        if not vals:
            return None, "string_mean", 0.0, 0.0
        if len(vals) == 1:
            w_e = 1.0 if iowa is not None else 0.0
            return float(vals[0]), "string_mean", w_e, 1.0 - w_e
        return float(sum(vals) / len(vals)), "string_mean", 0.5, 0.5
    return combine_empirical_and_modelled(iowa, orch, cfg)
```

**D. LaTeX:** production collections are IOWA and ORCH only. Philharmonia/McGill are excluded.

Strings:

$$
\mu(m)=
\begin{cases}
(y_I+y_O)/2 & \text{both present}\\
y_I\text{ or }y_O & \text{exactly one present}\\
\text{missing} & \text{neither}
\end{cases}
$$

Woodwind ordinario uses M-026 on the same pair. Example: $y_E=10$, $y_T=30$ $\Rightarrow$ `empirical_only` $10$, `equal_weight` $20$, `legacy_equal_weight` with $w_E=0.25$ $\Rightarrow$ $25$.

This is an **arithmetic** mean or convex combination, not a geometric mean.

**E. Symbols:** $y_I$ Iowa; $y_O$ Orchidea / transfer alias.

**F. Layman:** Strings: average Iowa and Orchidea when both exist. Woodwind ordinary (default): keep the real measurement and do not blend it with the family-transfer guess unless an equal-weight mode is selected.

**G. Specialist:** Arithmetic mean of a measured Orchidea effect and a **shared-$L$ Iowa model** is not two independent estimates (`SHARED_L_MEDIA_CAVEAT`). One present value is never halved. Zero is not used as a missing-value code. GUI Average and batch now share this resolver.

**H. Downstream:** in-memory Media, Zenodo Media, AcousticTable, summaries, Final_Calibration.

**I. Tests:** `test_export_consistency.py`; `test_media_of_prefer_measured_keeps_iowa` still covers the generic helper.

---

### M-028 — Preparatory `Media_*_mf` arithmetic pair average (offline)

**A. Status:** offline display column on the preparatory grid.

**B. Source:** [build_ste_preparatory.py](../build_ste_preparatory.py) line 948.

**C. Snippet:**

```python
        rec[f"Media_{spine}_mf"] = (ia + oa) / 2.0 if ia == ia and oa == oa else float("nan")
```

**D. LaTeX:** if both finite, $(y_I+y_O)/2$; else NaN. (`ia == ia` is a NaN self-inequality test.)

**E. Symbols:** mf ordinary/arco only.

**F. Layman:** Midpoint of Iowa and Orchidea at mf for the spine.

**G. Specialist:** Not `media_of`, not geometric, not empirical_only. Diagnostic column.

**H. Downstream:** preparatory Excel only.

**I. Tests:** No direct test identified.

---

### M-029 — Zenodo AcousticTable Media when layers are supplied (production)

**A. Status:** production path used by `export_zenodo_workbook` (layers is not `None`).

**B. Source:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) `_resolved_media` / `_acoustic_table_value` lines 415–425 and 948–953.

**C. Snippet:**

```python
    value, _, _, _ = resolve_iowa_orch_pair(
        iowa, orch, instrument=instrument, technique=technique, cfg=cfg
    )
    return value
```

When a woodwind `Final_Calibration` row exists, AcousticTable uses that row’s `final_value`.

**D. LaTeX:** same $\mu(m)$ as M-027. Example: string IOWA $=10$, ORCH $=30$ $\Rightarrow$ Media $=20$ in memory, Media sheet, and AcousticTable.

**E. Symbols:** `_layer_cdm` returns the cell value or `None`.

**F. Layman:** The published Media number is the same rule as the in-memory Media, not a second Iowa-first shortcut.

**G. Specialist:** Literals are written when layers are supplied. Reopen the workbook and read the stored number; do not treat formula text alone as a calculated value.

**H. Downstream:** AcousticTable numeric columns, gaps, $\ln$ Media.

**I. Tests:** `test_string_export_media_and_acoustic_are_twenty`; `test_woodwind_default_export_keeps_empirical_ten`; `test_acoustic_table_has_literals`.

---

### M-030 — Zenodo AcousticTable Media Excel formulas (fallback when `layers is None`)

**A. Status:** optional/fallback formula writer. Current `export_zenodo_workbook` passes layers, so this branch is unused in the standard runner.

**B. Source:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) `_media_sheet_formula` lines 389–412.

**C. Snippet:**

```python
        if method == "empirical_only":
            return f'=IF({iowa}<>"",{iowa},IF({orch}<>"",{orch},""))'
        if method == "legacy_equal_weight":
            we = float(cfg.empirical_weight)
            wt = 1.0 - we
            return (
                f'=IF(AND({iowa}<>"",{orch}<>""),{we}*{iowa}+{wt}*{orch},'
                f'IF({iowa}<>"",{iowa},IF({orch}<>"",{orch},"")))'
            )
    return f'=IF(COUNT({iowa}:{orch})=0,"",AVERAGE({iowa}:{orch}))'
```

**D. LaTeX:** strings and `equal_weight` use Excel `AVERAGE` of the non-blank pair (one present $\Rightarrow$ that one). Woodwind `empirical_only` uses Iowa-else-Orch. `legacy_equal_weight` writes $w_E y_I+w_T y_O$. Delegated: **L-007**.

**E. Symbols:** sheet cells.

**F. Layman:** Excel’s ordinary average of the two collection columns.

**G. Specialist:** Matches the comment at line 871 (“never halved when only one exists”) and string `media_of`, not M-029.

**H. Downstream:** only if someone exports without in-memory layers.

**I. Tests:** `test_media_formulas_quote_bass_clarinet_sheets` (quoting), not numerical AVERAGE.

---

### M-031 — Origin-rank overwrite (production)

**A. Status:** production for `Layer.place` default `prefer_measured`.

**B. Source:** [ste_lab/catalog.py](../ste_lab/catalog.py) `ORIGIN_RANK` lines 194–214; [ste_lab/session.py](../ste_lab/session.py) lines 121–128.

**C. Snippet:**

```python
        if origin_rank(existing.origin) <= origin_rank(origin):
            if existing.origin != origin and abs(existing.value - float(value)) > 1e-9:
                existing.flags.append("duplicate_conflict_kept")
            return "merged_kept"
```

**D. LaTeX:** keep existing if $\mathrm{rank}(\mathrm{old})\le\mathrm{rank}(\mathrm{new})$ (lower rank wins). Absolute difference threshold $10^{-9}$ only flags; it does not average.

**E. Symbols:** ranks: measured $0$ … polynomial/extrapolated $8$.

**F. Layman:** A measured number is not overwritten by a modelled one unless the caller forces overwrite.

**G. Specialist:** Discrete priority, not a Kalman update. Batch `build_layer` uses `overwrite="overwrite"`, so this rank rule is skipped on first construction.

**H. Downstream:** GUI paste.

**I. Tests:** No direct test identified.

---

### Stage F — Missing-note fill

### M-032 — `fill_missing` interpolation / capped extrapolation (optional GUI; not woodwind two-ratio)

**A. Status:** optional. Default method `pchip`. Woodwind two-ratio does **not** call this (methodology sheet: “Not used in the two-ratio transfer”).

**B. Source:** [ste_lab/transfer.py](../ste_lab/transfer.py) lines 235–308. GUI: [ste_lab/gui_app.py](../ste_lab/gui_app.py) `_run_fill` (catches `InsufficientFillData`; layer unchanged).

**C. Snippet:**

```python
    if method == "pchip":
        if len(known) < 3:
            raise InsufficientFillData(
                f"PCHIP fill needs at least 3 known values; this layer has {len(known)}. "
                "Choose linear fill for two points, or add another measured note. "
                "No other method is substituted automatically."
            )
        fn = PchipInterpolator(xs, ys, extrapolate=True)
        extra_origin = "extrapolated_pchip"
    elif method == "linear":
        extra_origin = "extrapolated_hold"
```

Default `max_extrap_semitones=12`, `poly_degree=3`. Linear/polynomial require $\ge 2$ known cells. PCHIP requires $\ge 3$ or raises `InsufficientFillData` (not `Unknown fill method`). The GUI keeps the layer unchanged and recommends linear.

**D. LaTeX:** known $(m_j,y_j)$, span $[m_{\min},m_{\max}]$. Fill $m\in[\max(\ell,m_{\min}-c),\min(h,m_{\max}+c)]$ missing from the layer:

$$
\hat y(m)=
\begin{cases}
\mathrm{PCHIP}(m;\{m_j,y_j\}) & \texttt{pchip},\ n\ge 3\ \text{(extrapolate=True)}\\
\mathrm{np.interp}(m) & \texttt{linear}\\
\sum_{k=0}^{d} a_k m^k & \texttt{polynomial},\ d=\min(d_{\mathrm{req}},n-1),\ d\ge 1
\end{cases}
$$

drop if $\hat y\le 0$. Interior origin `interpolated`. Exterior origins: `extrapolated_pchip`, `extrapolated_hold`, or `extrapolated_polynomial`. Historical `extrapolated_ridge` tags remain readable.

**E. Symbols:** $y$ in **value** space, not $\ln y$. $c$ cap in semitones. PCHIP/polyfit: L-001, L-003. `np.interp` outside the span **holds edges** (NumPy).

**F. Layman:** Fill empty notes. Inside the measured stretch, draw a smooth (or straight) line. Outside, you may continue a short way, depending on the method. Fill is **not** interior-only.

**G. Specialist:** **Disagrees with M-016:** here PCHIP **does** cubic-extrapolate (`extrapolate=True`). That edge policy was **not** changed to match technique-transfer hold-$L$. There is **no ridge penalty**; the old `extrapolated_ridge` tag was a mislabel. Polynomial is global least squares on raw $y$ vs MIDI (not log $y$). QA flags degree $\ge 3$ and $c>12$. New origins are in `GENERATED_ORIGINS` / `ORIGIN_TO_PROVENANCE`.

**H. Downstream:** GUI layers only unless a script calls `fill_missing`.

**I. Tests:** `TestFillHandling` in `test_export_consistency.py`.

---

### Stage G — Validation and disagreement

### M-033 — Finite overlap pairs (production)

**A. Status:** production helper.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `finite_pairs`, lines 337–343.

**C. Snippet:**

```python
        if va and vb and va > 0 and vb > 0 and math.isfinite(va) and math.isfinite(vb):
            out.append((int(midi), float(va), float(vb)))
```

**D. LaTeX:** pairs $(m,y_a,y_b)$ with both values finite and $>0$. Note: `if va and vb` also rejects exact $0.0$ (already excluded) and would reject a hypothetical `NaN` via `isfinite`.

**E. Symbols:** as M-010.

**F. Layman:** Only notes where both curves have a usable number.

**G. Specialist:** Truth and prediction must share MIDI. No interpolation of the validator.

**H. Downstream:** M-035–M-037.

**I. Tests:** empty/NaN test.

---

### M-034 — Extrapolation distance in semitones (production)

**A. Status:** production.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py) lines 161–175.

**C. Snippet:**

```python
def measured_span(curve: dict[int, float]) -> tuple[Optional[int], Optional[int]]:
    keys = [m for m, v in curve.items() if v is not None and v > 0 and math.isfinite(v)]
    if not keys:
        return None, None
    return min(keys), max(keys)

def extrapolation_distance_semitones(midi: int, lo: Optional[int], hi: Optional[int]) -> int:
    if lo is None or hi is None:
        return 0
    if lo <= midi <= hi:
        return 0
    if midi < lo:
        return lo - midi
    return midi - hi
```

**D. LaTeX:** span of **empirical** positive keys $[\ell,h]$.

$$
d(m)=
\begin{cases}
0 & \ell\text{ or }h\text{ missing, or }\ell\le m\le h\\
\ell-m & m<\ell\\
m-h & m>h
\end{cases}
$$

**E. Symbols:** $d$ in semitones (MIDI difference). Span is Iowa empirical in `build_final_rows`, not the transfer-anchor span.

**F. Layman:** How many notes this pitch sits beyond the last Iowa measurement.

**G. Specialist:** Distance to the empirical min/max, not to the nearest measured neighbour (interior holes have $d=0$). Missing span $\Rightarrow d=0$ (no penalty).

**H. Downstream:** M-041 provenance `TRANSFERRED_EXTRAPOLATED` when $d>0$.

**I. Tests:** `test_distances`; `test_long_extrapolation_rejected`.

---

### M-035 — Agreement metrics (production validation)

**A. Status:** production validation tables (woodwind). Not used to retune default `empirical_only`.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `agreement_metrics`, lines 346–375.

**C. Snippet:**

```python
    err = [p - t for _, p, t in pairs]
    abs_err = [abs(e) for e in err]
    mae = sum(abs_err) / n
    rmse = math.sqrt(sum(e * e for e in err) / n)
    ln_err = [math.log(p) - math.log(t) for _, p, t in pairs]
    ln_rmse = math.sqrt(sum(e * e for e in ln_err) / n)
    medae = sorted(abs_err)[n // 2]
    mape = sum(abs(p - t) / t for _, p, t in pairs) / n
    r = None
    if n >= 3:
        mp = sum(p for _, p, _ in pairs) / n
        mt = sum(t for _, _, t in pairs) / n
        num = sum((p - mp) * (t - mt) for _, p, t in pairs)
        denp = math.sqrt(sum((p - mp) ** 2 for _, p, _ in pairs))
        dent = math.sqrt(sum((t - mt) ** 2 for _, _, t in pairs))
        if denp > 0 and dent > 0:
            r = num / (denp * dent)
```

**D. LaTeX:** $n$ overlap pairs $(\hat y_i,y_i)$.

$$
\mathrm{MAE}=\frac1n\sum|\hat y_i-y_i|,\quad
\mathrm{RMSE}=\sqrt{\frac1n\sum(\hat y_i-y_i)^2}
$$

$$
\ln\mathrm{RMSE}=\sqrt{\frac1n\sum\bigl(\ln\hat y_i-\ln y_i\bigr)^2}
$$

$$
\mathrm{MedAE}=|e|_{(n//2+1)}\quad\text{(0-based }n//2\text{; not averaged for even }n\text{)}
$$

$$
\mathrm{MAPE}=\frac1n\sum\frac{|\hat y_i-y_i|}{y_i}
$$

$$
r=\frac{\sum(\hat y_i-\bar{\hat y})(y_i-\bar y)}{\sqrt{\sum(\hat y_i-\bar{\hat y})^2}\sqrt{\sum(y_i-\bar y)^2}}\quad(n\ge 3;\ \text{else }r=\emptyset)
$$

Docstring states $\ln$RMSE is RMSE in log space, **not** MSLE.

**E. Symbols:** $\hat y$ prediction (Iowa, transfer, or “combined_empirical_only” which is **Iowa again** in `build_validation_tables` line 689). $y$ validator PHIL or McGill. $r$ is Pearson on **raw** $y$, not on $\ln y$.

**F. Layman:** Several ways of saying how close two collections are on the notes they share.

**G. Specialist:** Validators are different collections/codecs; agreement is **not** a hold-out of the transfer fit (Iowa is excluded as truth, but PHIL/McGill are not independent of instrument physics). Population ($/n$) RMSE. MedAE is a lower/upper-middle order statistic, not M-022. MAPE is not multiplied by 100. `combined_empirical_only` duplicates the Iowa-vs-validator row.

**H. Downstream:** Validation sheet; copied onto `build_final_rows` as optional columns (callers may leave them `None` — `excel_export` does not pass per-row validation scores into `build_final_rows`).

**I. Tests:** `test_ln_rmse_applies_square_root`; `test_holdout_ln_rmse_order`.

---

### M-036 — Shape-normalised log residuals (validation only)

**A. Status:** research/validation diagnostic. Explicitly does not modify production values.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `shape_normalised_metrics`, lines 388–435.

**C. Snippet:**

```python
        logs.append((int(midi), math.log(float(y)), math.log(float(y_hat))))
```

```python
    c = _median([ly - lh for _, ly, lh in logs])
```

```python
    residuals = [ly - (lh + c) for _, ly, lh in logs]
```

```python
        "collection_log_offset": c,
        "collection_scale_factor": math.exp(c),
        "shape_ln_rmse": math.sqrt(sum(e * e for e in residuals) / n),
        "shape_ln_mae": sum(abs(e) for e in residuals) / n,
```

**D. LaTeX:**

$$
c=\mathrm{median}_i\bigl(\ln y_i-\ln\hat y_i\bigr),\qquad
s=e^{c}
$$

$$
e_i=\ln y_i-(\ln\hat y_i+c),\qquad
\mathrm{shape\_ln\_RMSE}=\sqrt{\frac1n\sum e_i^2},\quad
\mathrm{shape\_ln\_MAE}=\frac1n\sum|e_i|
$$

Skip non-positive/non-finite $y$ or $\hat y$.

**E. Symbols:** $c$ is a **log-offset**, not dB. $s$ would multiply $\hat y$ to match the typical level of $y$. **Not applied** to production $\hat y$.

**F. Layman:** After ignoring a constant “this collection is generally louder/denser,” how well do the shapes match?

**G. Specialist:** Median log residual is a robust collection scale. Residual definition uses observed minus (predicted $+c$). This is **not** a mixed-effects model and does not estimate uncertainty of $c$.

**H. Downstream:** validation export only.

**I. Tests:** `test_shape_metrics_analytical_fixture`; `test_shape_does_not_change_production_values`.

---

### M-037 — Collection offset as median ratio (validation)

**A. Status:** validation.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `collection_offset`, lines 438–447.

**C. Snippet:**

```python
    """Median a/b on overlap — collection offset, not a transfer quality score."""
    ratios = [p / t for _, p, t in finite_pairs(a, b)]
```

**D. LaTeX:** $\mathrm{median}(y_a/y_b)$ via M-022.

**E. Symbols:** dimensionless. Caller uses Iowa vs PHIL/McGill.

**F. Layman:** Typical Iowa-to-other-library ratio.

**G. Specialist:** Location shift on the **ratio** scale, not a correlation. Docstring: not a transfer quality score.

**H. Downstream:** Offsets sheet.

**I. Tests:** No dedicated assertion beyond table construction.

---

### M-038 — Grid-search convex weight (research / unused in production)

**A. Status:** research/test-only. Implemented and unit-tested; **no production caller**.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `optimal_weight`, lines 450–480.

**C. Snippet:**

```python
    w = 0.0
    while w <= 1.0 + 1e-12:
        pred = {m: w * empirical[m] + (1.0 - w) * transfer[m] for m in midis}
        mets = agreement_metrics(pred, {m: truth[m] for m in midis})
        score = mets["mae"] if metric == "mae" else mets["rmse"]
        if score is not None and (best_score is None or score < best_score):
            best_w, best_score, best_metrics = w, score, mets
        w = round(w + step, 6)
```

Default `step=0.01`, `min_n=5`, `metric="mae"`.

**D. LaTeX:** $\hat y(w)=w y_E+(1-w)y_T$, $w\in\{0,0.01,\ldots,1\}$. Choose $w$ minimising MAE (or RMSE). Requires overlap of empirical, transfer, and truth with $n\ge n_{\min}$.

**E. Symbols:** $w$ dimensionless. Truth would be PHIL/McGill if it were wired.

**F. Layman:** Try many mix percentages and pick the closest to a third library.

**G. Specialist:** In-sample tuning on the validator; would leak if used as production. YAML `validation_optimised` does **not** invoke this.

**H. Downstream:** none in runners.

**I. Tests:** `test_obvious_weight`; `test_too_small_overlap_returns_none`.

---

### M-039 — Absolute and relative disagreement (production)

**A. Status:** production on final-calibration rows.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `disagreement`, lines 483–490. Default `DEFAULT_REL_EPS = 1e-12` (line 38); YAML `disagreement.relative_epsilon`.

**C. Snippet:**

```python
    delta = abs(empirical - transfer)
    rel = delta / max(abs(empirical), eps)
    return delta, rel
```

Returns `(None, None)` if either value is missing or non-finite.

**D. LaTeX:**

$$\Delta=|y_E-y_T|,\qquad \delta_{\mathrm{rel}}=\frac{\Delta}{\max(|y_E|,\varepsilon)},\quad \varepsilon=10^{-12}$$

**E. Symbols:** same metric units for $\Delta$; $\delta_{\mathrm{rel}}$ dimensionless. Denominator uses $|y_E|$, not $\max(|y_E|,|y_T|)$.

**F. Layman:** How far the measurement and the model sit apart, as a gap and as a fraction of the measurement.

**G. Specialist:** Relative scale is one-sided (empirical). QA uses $\delta_{\mathrm{rel}}>0.5$ only for `COMBINED_ESTIMATE` (M-041). Not a statistical test.

**H. Downstream:** Final_Calibration columns; QA review on combined estimates.

**I. Tests:** No isolated disagreement unit test identified.

---

### Stage H — Provenance, QA, and acceptance

### M-040 — Provenance of a final cell (production)

**A. Status:** production in `build_final_rows`. Related mapper `provenance_of_origin` (lines 150–158) is used for Measured_Data stamps, not this block.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py) lines 739–761.

**C. Snippet:**

```python
        if e is not None:
            prov = PROVENANCE_MEASURED
        elif t is not None and transfer_method == "legacy_target_calibrated":
            prov = PROVENANCE_COMBINED_ESTIMATE
        elif t is not None and outside:
            prov = PROVENANCE_TRANSFERRED_EXTRAPOLATED
        elif t is not None:
            prov = PROVENANCE_TRANSFERRED
        else:
            continue
```

`outside` is $d(m)>0$ from M-034 on the **empirical** span.

**D. LaTeX:** piecewise label, not a numeric estimator. A positive finite $y_E$ is always `MEASURED`, even if $y_T$ exists.

**E. Symbols:** labels are strings. `INTERPOLATED` is in `ORIGIN_TO_PROVENANCE` but is not assigned in this block.

**F. Layman:** Stamp saying whether the published number is a measurement, a transfer, a transfer beyond Iowa’s range, or a deprecated blend.

**G. Specialist:** “MEASURED” means “empirical argument present,” i.e. a positive Iowa (or mapped empirical) value in `build_final_rows`, not an independent audit of the recording. Legacy scalar cannot be `TRANSFERRED` by construction.

**H. Downstream:** M-041; AcousticTable when `final_by_cell` is attached.

**I. Tests:** `test_measured_stays_measured`; `test_transferred_inside_span`; `test_transferred_extrapolated_outside_span`; `test_legacy_scalar_is_combined_estimate`.

---

### M-041 — QA status piecewise rule (production)

**A. Status:** production. Defaults from YAML: accept $\le 3$ st, review $\le 6$ st.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `qa_status`, lines 493–525.

**C. Snippet:**

```python
    if value is None or not math.isfinite(value) or value <= 0:
        return QA_REJECTED
    if provenance == PROVENANCE_MEASURED:
        return QA_HIGH
    if provenance == PROVENANCE_INTERPOLATED and extrap_st == 0:
        return QA_HIGH
    if extrap_st > cfg.review_max_semitones:
        return QA_REJECTED if not cfg.allow_long_extrapolation else QA_LOW
    if extrap_st > cfg.accept_max_semitones:
        return QA_REVIEW
    if provenance == PROVENANCE_TRANSFERRED_EXTRAPOLATED:
        return QA_REVIEW if extrap_st > cfg.accept_max_semitones else QA_MODERATE
    if provenance == PROVENANCE_TRANSFERRED:
        return QA_MODERATE
    if provenance == PROVENANCE_COMBINED_ESTIMATE:
        if relative_diff is not None and relative_diff > 0.5:
            return QA_REVIEW
        return QA_MODERATE
    if has_empirical:
        return QA_HIGH
    if has_transfer:
        return QA_MODERATE
    return QA_LOW
```

**D. LaTeX:** let $d$ be M-034, $A=3$, $R=6$ by default. Sequential rules:

- $\hat y\le 0$ or non-finite $\Rightarrow$ REJECTED
- MEASURED, or INTERPOLATED with $d=0$ $\Rightarrow$ HIGH
- $d>R$ and long-extrap forbidden $\Rightarrow$ REJECTED; if allowed $\Rightarrow$ LOW
- $A<d\le R$ $\Rightarrow$ REVIEW_REQUIRED
- TRANSFERRED_EXTRAPOLATED: REVIEW if $d>A$, else MODERATE
- TRANSFERRED $\Rightarrow$ MODERATE
- COMBINED_ESTIMATE: REVIEW if $\delta_{\mathrm{rel}}>0.5$, else MODERATE
- else HIGH / MODERATE / LOW fallbacks

A measured cell never reaches the $d>R$ reject branch.

**E. Symbols:** $d$ semitones; $\delta_{\mathrm{rel}}$ from M-039. Threshold $0.5$ is hard-coded, not YAML.

**F. Layman:** A traffic-light for whether the number is trusted enough to publish.

**G. Specialist:** Policy, not a p-value. HIGH for every measurement including outliers. REVIEW is descriptive; production eligibility is M-042.

**H. Downstream:** M-042; AcousticTable `accepted_final`.

**I. Tests:** `test_high_measured_is_accepted`; `test_moderate_short_extrap_is_accepted`; `test_review_required_not_accepted_by_default`; `test_rejected_not_accepted`; `test_review_override_preserves_qa_flag`.

---

### M-042 — Automatic and final acceptance (production)

**A. Status:** production. Default: HIGH and MODERATE accepted; REVIEW and REJECTED not accepted.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py) lines 528–554.

**C. Snippet:**

```python
def automatic_acceptance(status: str, cfg: CalibrationConfig) -> bool:
    if status in {QA_HIGH, QA_MODERATE}:
        return True
    if status == QA_LOW and cfg.allow_long_extrapolation:
        return True
    return False

def accepted_for_final(status: str, cfg: CalibrationConfig) -> bool:
    if automatic_acceptance(status, cfg):
        return True
    if status == QA_REVIEW and cfg.allow_review_required:
        return True
    return False
```

Rejected cells keep `final_value=None` (`keep_value = final if status != QA_REJECTED else None`).

**D. LaTeX:** Boolean policy. `allow_review_required` can accept REVIEW while **preserving** the QA flag (`acceptance_override`).

**E. Symbols:** none.

**F. Layman:** Only green/amber lights are published unless you flip a YAML override.

**G. Specialist:** Not a multiple-testing correction. REVIEW visible-but-ineligible is the default 1.3.x hygiene choice.

**H. Downstream:** `validation_status_of`; `acceptance_counts`.

**I. Tests:** same as M-041; `test_acceptance_counts_reconcile`; `test_acoustic_table_follows_final_calibration`.

---

### M-043 — Layer $3.5\sigma$ outlier flag (advisory QA)

**A. Status:** advisory (does not change values).

**B. Source:** [ste_lab/qa.py](../ste_lab/qa.py) lines 196–215.

**C. Snippet:**

```python
    values = [c.value for c in cells if 0 < c.value < 1e6]
    if len(values) >= 6:
        mean = sum(values) / len(values)
        var = sum((v - mean) ** 2 for v in values) / len(values)
        sd = var ** 0.5
        if sd > 0:
            for cell in cells:
                if abs(cell.value - mean) > 3.5 * sd:
```

**D. LaTeX:** on values $0<y<10^6$, $n\ge 6$:

$$\bar y=\frac1n\sum y_i,\quad
\sigma=\sqrt{\frac1n\sum(y_i-\bar y)^2}
\quad\text{(population SD, divisor }n\text{)}$$

flag if $|y-\bar y|>3.5\sigma$.

**E. Symbols:** $y$ raw metric, not $\ln y$. Threshold $10^6$ drops pathological pastes.

**F. Layman:** Mark notes that sit far from the layer’s average.

**G. Specialist:** Heuristic. Population variance; no Student-$t$; no robust MAD. Mixing measured and modelled cells in one layer inflates or shrinks $\sigma$. Not used by M-041.

**H. Downstream:** QA_Flags sheet only.

**I. Tests:** No direct test identified.

---

### M-044 — Sparse-measured share and register-span flags (advisory QA)

**A. Status:** advisory.

**B. Source:** [ste_lab/qa.py](../ste_lab/qa.py) lines 49–98.

**C. Snippet:**

```python
    measured_share = len(measured) / n
    if measured_share < 0.25 and n >= 8:
```

```python
        span = max(midis) - min(midis)
        if measured:
            mspan = max(c.midi for c in measured) - min(c.midi for c in measured)
            beyond = max(
                max(midis) - max(c.midi for c in measured),
                min(c.midi for c in measured) - min(midis),
            )
            if beyond > 12:
```

```python
            if mspan < 12 and span > 24:
```

**D. LaTeX:** $p_{\mathrm{meas}}=n_{\mathrm{meas}}/n$. Flag if $p_{\mathrm{meas}}<0.25$ and $n\ge 8$. Let $s$ be full MIDI span, $s_M$ measured span, $b$ max extension beyond measured extrema. Flag $b>12$; critical if $s_M<12$ and $s>24$.

**E. Symbols:** MIDI differences in semitones.

**F. Layman:** Warns when most of the curve is invented or stretched far past real notes.

**G. Specialist:** Thresholds are workbook-history heuristics (viola tail), not theoretically derived.

**H. Downstream:** QA_Flags.

**I. Tests:** No direct test identified.

---

### M-045 — Modelled dynamic inversion count (advisory QA)

**A. Status:** advisory.

**B. Source:** [ste_lab/qa.py](../ste_lab/qa.py) lines 354–374.

**C. Snippet:**

```python
                if b.value + 0.25 < a.value:
                    inversions += 1
                    if (a.origin.lower() in GENERATED_ORIGINS) and (b.origin.lower() in GENERATED_ORIGINS):
                        both_modelled += 1
            if both_modelled >= 4:
```

**D. LaTeX:** count MIDI where $y_{\mathrm{ff}}+0.25<y_{\mathrm{pp}}$ and both origins are generated. Flag if that count $\ge 4$.

**E. Symbols:** $0.25$ is in **raw metric units**, not dB. Meaning depends on F-061 scale (unknown here).

**F. Layman:** If the loud dynamic is repeatedly smaller than the soft one on invented notes, raise a flag.

**G. Specialist:** Not a monotone-loudness test on measurements. The $0.25$ slack is a heuristic.

**H. Downstream:** QA_Flags.

**I. Tests:** No direct test identified.

---

### Stage I — Summaries, AcousticTable derived fields, evidence

### M-046 — Geometric mean of positive values (production summary)

**A. Status:** production on Summary / Empirical_ORCH summary. **Not** the Media point estimator (that is M-027).

**B. Source:** [ste_lab/excel_export.py](../ste_lab/excel_export.py) lines 342–348; [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) lines 770–774 and 1082.

**C. Snippet:**

```python
        logs = np.log(vals)
        gmean = float(np.exp(np.mean(logs)))
```

```python
    def gmean(vals):
        vals = [v for v in vals if v and v > 0]
        if not vals:
            return None
        return math.exp(sum(math.log(v) for v in vals) / len(vals))
```

Excel fallback: `=GEOMEAN(...)` (L-005).

**D. LaTeX:**

$$G=\exp\Bigl(\frac1n\sum_{i=1}^n \ln y_i\Bigr)=\Bigl(\prod_{i=1}^n y_i\Bigr)^{1/n},\quad y_i>0$$

Ratio_extrapol Summary uses Media-layer cells after dropping `above_measured_ceiling` and $y\le 0$. Empirical_ORCH summary uses **measured principal rows only**.

**E. Symbols:** $y$ F-061-like. $G$ has the same units as $y$. Natural log.

**F. Layman:** A typical density that is not pulled as hard by very large notes as an ordinary average would be.

**G. Specialist:** Back-transformed arithmetic mean of $\ln y$. Not a median. Woodwind Empirical_ORCH $G$ is the inferential “typical measured cell”; string Media $G$ averages a completed grid and is weaker evidence (`SHARED_L_MEDIA_CAVEAT`).

**H. Downstream:** Summary_Measured_Range; Zenodo Summary.

**I. Tests:** No numerical GEOMEAN fixture identified.

---

### M-047 — Registral log-slope and octave change (production summary)

**A. Status:** production on the Ratio_extrapol Summary when $n\ge 3$. Zenodo formula path uses Excel `SLOPE`/`EXP` (L-006, L-008).

**B. Source:** [ste_lab/excel_export.py](../ste_lab/excel_export.py) lines 349–353.

**C. Snippet:**

```python
        slope = float(np.polyfit(midis, logs, 1)[0])
        oct_chg = float(math.exp(slope * 12) - 1)
        corr = float(np.corrcoef(midis, logs)[0, 1])
```

**D. LaTeX:** OLS $\ln y = a + b\, m +\varepsilon$ (L-003). Report $b$ (ln per semitone),

$$\Delta_{\mathrm{oct}}=e^{12b}-1$$

and Pearson $r(m,\ln y)$ (L-004).

**E. Symbols:** $m$ MIDI; $b$ in $\ln(y)$ per semitone; $\Delta_{\mathrm{oct}}$ relative change per 12 semitones. Not dB/octave.

**F. Layman:** Does density tend to rise or fall as notes get higher, and by what percent over an octave?

**G. Specialist:** Linear log-density vs MIDI is a descriptive fit, not a source-filter model. Same exclusion rules as M-046.

**H. Downstream:** Summary sheet only.

**I. Tests:** No direct test identified.

---

### M-048 — AcousticTable gaps, half-gaps, and $\ln$ Media (production)

**A. Status:** production derived columns.

**B. Source:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) lines 436–494 (Python) and 514–529 (Excel).

**C. Snippet:**

```python
    def _sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
        if a is None or b is None:
            return None
        return a - b

    def _half(a: Optional[float], b: Optional[float]) -> Optional[float]:
        if a is None or b is None:
            return None
        return abs(a - b) / 2.0

    def _ln(v: Optional[float]) -> Optional[float]:
        if v is None or v <= 0:
            return None
        return math.log(v)
```

**D. LaTeX:** examples at a MIDI:

$$
g_{\mathrm{mf-pp}}^{\mathrm{IOWA}}=y_{\mathrm{mf}}^I-y_{\mathrm{pp}}^I,\quad
h_{\mathrm{pp}}=\tfrac12\bigl|y_{\mathrm{pp}}^{O}-y_{\mathrm{pp}}^{I}\bigr|,\quad
\ell_{\mathrm{pp}}=\ln y_{\mathrm{Media,pp}}
$$

when operands exist and, for $\ln$, $y>0$.

**E. Symbols:** raw-metric differences (not log-ratios). Half-gap is a **dispersion** between collections, not a standard error.

**F. Layman:** Extra columns showing how dynamics differ and how far Iowa and Orchidea sit apart, plus the log of the Media number.

**G. Specialist:** Arithmetic gaps on an unknown-scale metric are not dynamic “intervals” in dB. Media definition is M-029 or M-030 depending on branch.

**H. Downstream:** display/chart; Excel `SLOPE` uses $\ln$ columns.

**I. Tests:** `test_acoustic_table_has_literals`.

---

### M-049 — Evidence-role classification (production policy; affects which $y$ enter Empirical_ORCH)

**A. Status:** production labelling. Not a numeric estimator, but it **selects** the inferential dataset.

**B. Source:** [ste_lab/evidence.py](../ste_lab/evidence.py), `evidence_role` / `stamp_evidence` / `principal_empirical_rows`, lines 62–207.

**C. Snippet (rules):** prediction techniques (`sul tasto`); context collections PHIL/McGill; inherited dynamic if `l_donor_dynamic` $\neq$ layer dynamic; IOWA effects $\to$ `modelled_completion`; ORCH measured $\to$ `empirical`; woodwind Iowa ordinario measured $\to$ `empirical` if no ORCH principal rows.

**D. LaTeX:** discrete classifier. Principal rows $P$: ORCH measured cells if any; else woodwind non-effect Iowa measured cells; else $\emptyset$.

**E. Symbols:** none.

**F. Layman:** The program decides which numbers count as “real evidence” versus “completed grid” or “context.”

**G. Specialist:** Shared-$L$ Iowa effects are excluded from principal evidence by design. From 1.5.4, a technique missing from IOWA/ORCH may be taught by an extra collection’s $L=\ln(y/x)$. A missing production dynamic $d$ uses the closest same-collection teacher $d_*$ on $\mathrm{pp},p,\mathrm{mp},\mathrm{mf},f,\mathrm{ff}$: $\hat y(d)=z_{\mathrm{ord}}(d)\,e^{L(d_*)}$. `l_donor_dynamic` records $d_*$. If two or more teacher dynamics exist, `l_invariance` is `held` when $\max\bar L-\min\bar L\le 0.40$ nats, else `wide`; a single teacher is `untested`. A foreign leftover (Phil $p$ when Orchidea already teaches $mf$) is not used. A nested Philharmonia `arco-sul-tasto` book is sul tasto, never ordinario. This is a research-design rule, not a likelihood.

**H. Downstream:** Empirical_ORCH sheet; M-046 principal $G$.

**I. Tests:** `test_orch_recorded_is_principal`; `test_iowa_ordinario_is_principal_on_woodwind`; `test_family_transfer_orch_is_not_principal`; `test_mcgill_and_phil_are_context_not_principal`.

---

### M-050 — YAML / simple mapping parse (configuration, not a scientific estimator)

**A. Status:** production config loader.

**B. Source:** [ste_lab/calibration.py](../ste_lab/calibration.py), `_parse_simple_yaml`, lines 113–147.

**C. Snippet:** types `true`/`false`, `int`, then `float`, else string. Comments stripped at `#`.

**D. LaTeX:** none. Defaults listed in `CalibrationConfig` (lines 60–74) if the file is missing.

**E. Symbols:** `seed` is stored (`seed: int = 0`) and **never read** by a calculation.

**F. Layman:** Reads the settings file without installing PyYAML.

**G. Specialist:** Nested lists unsupported. `transfer_weight` is parsed and must equal $1-w_E$ (M-026).

**H. Downstream:** all calibration numerics.

**I. Tests:** `test_config_loads`.

---

## External-library operations

Project algorithms that merely call `log`/`exp`/`mean` are documented as M-entries. The following are **delegated** solvers or spreadsheet functions. Internal equations of those libraries are not reproduced.

### L-001 — `scipy.interpolate.PchipInterpolator` (production / optional)

**Package:** SciPy `scipy` 1.13.1 in the audit environment (requirement `scipy>=1.11`).  
**Import:** `from scipy.interpolate import PchipInterpolator`.

**Call sites:**

1. [ste_lab/transfer.py](../ste_lab/transfer.py) line 66 — `_interp_log_ratio`, `PchipInterpolator(xs, ys, extrapolate=False)` on **log-ratios**, $n\ge 3$.
2. [ste_lab/transfer.py](../ste_lab/transfer.py) line 253 — `fill_missing`, `PchipInterpolator(xs, ys, extrapolate=True)` on **raw values**, $n\ge 3$.

**Arguments:** monotone-increasing unique `xs` (MIDI); `ys` corresponding $L$ or $y$. Default `axis=0`.

**Default-dependent behaviour (SciPy 1.13 docstring, local inspect):** piecewise cubic Hermite interpolant constructed to be shape-preserving / monotonic in each interval; `extrapolate=False` returns NaNs outside (call site 1 never evaluates outside; it holds edges itself). `extrapolate=True` continues using the first/last interval polynomial (call site 2).

**Layman:** A smooth curve used to fill $L$ or $y$ between known notes, with fewer extra wiggles than a single high-degree polynomial.

**Specialist:** Fritsch–Carlson-type PCHIP. Not a statistical smoother; no confidence band. Duplicate $x$ is illegal. Call site 2’s extrapolation is the behaviour M-016 was written to avoid on $L$.

**I/O:** $x$: MIDI; $y$: dimensionless $L$ or F-061-like $y$.

---

### L-002 — `numpy.interp` (production / optional)

**Package:** NumPy.  
**Call sites:** [ste_lab/transfer.py](../ste_lab/transfer.py) line 64 (`n=2` log-ratio interior); line 261 (`fill_missing` linear).

**Arguments:** `np.interp(m, xs, ys)` with default `left`/`right` = end values.

**Default-dependent behaviour:** piecewise linear; **constant hold** outside `[xs[0], xs[-1]]`. In `_interp_log_ratio` the project already restricts this call to the interior.

**Layman:** Straight line between two (or many) points.

**Specialist:** Linear interpolation of $L$ vs MIDI is linear in log-ratio, i.e. a power-law blend of ratios, not of $y$.

---

### L-003 — `numpy.polyfit` / `numpy.polyval` (optional fill; summary slope)

**Call sites:** [ste_lab/transfer.py](../ste_lab/transfer.py) lines 266–269; [ste_lab/excel_export.py](../ste_lab/excel_export.py) line 351 (`np.polyfit(midis, logs, 1)[0]`).

**Project model:** polynomial $y\approx\sum_{k=0}^{d}a_k m^k$ or $\ln y\approx a+bm$. Degree $d=\max(1,\min(d_{\mathrm{req}},n-1))$.

**Library role:** least-squares coefficient solver. Internals not documented here.

**Layman:** Best-fit polynomial (or straight line on the logs).

**Specialist:** Global polynomial on a short measured core is the failure mode QA `polynomial_underdetermined` describes. No regularisation (the origin name `extrapolated_ridge` is misleading).

---

### L-004 — `numpy.corrcoef` (production summary)

**Call site:** [ste_lab/excel_export.py](../ste_lab/excel_export.py) line 353, `np.corrcoef(midis, logs)[0, 1]`.

**Project meaning:** Pearson $r$ between MIDI and $\ln y$ (M-047).

**Library role:** sample correlation matrix. Defaults: no weights.

---

### L-005 — Excel `GEOMEAN` (exported formula)

**Call site:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) lines 836–838.

**Project meaning:** same target as M-046 on Media columns D/G/J through the last measured-range row.

**Library role:** Excel geometric mean. Not evaluated in this audit. Requires positive numbers.

---

### L-006 — Excel `SLOPE` (exported formula)

**Call site:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) lines 842–844, `SLOPE(ln Media, MIDI)`.

**Project meaning:** analogue of $b$ in M-047.

**Library role:** Excel OLS slope. Internals not reproduced.

---

### L-007 — Excel `AVERAGE` (exported formula)

**Call site:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) `_media_sheet_formula` lines 389–412.

**Project meaning:** M-030 Media. Arithmetic mean of numeric cells; one of two blank $\Rightarrow$ the remaining value (`COUNT` guard).

---

### L-008 — Excel `LN`, `EXP`, `CORREL`, `COUNT` (exported formulas)

**Call sites:** [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) lines 527–529 (`LN`), 847–849 (`EXP(12*slope)-1`), 854–856 (`CORREL`), 860–862 (`COUNT`).

**Project meaning:** natural log of Media; octave relative change; Pearson of $\ln y$ vs MIDI; counts. Excel `LN` is natural log, consistent with `math.log`.

---

### L-009 — `pandas.read_excel` / `ExcelFile` (ingest, not a formula)

**Call sites:** preparatory and batch loaders (`load_preparatory`, `load_spectral_mass`, `build_ste_preparatory`).

**Project role:** cell values become floats; header hunt is string matching. No statistical transform. Type coercion can drop rows (missingness-by-parse, not imputation).

---

### L-010 — `numpy.median` (optional GUI family/linear ratio)

**Call sites:** [ste_lab/transfer.py](../ste_lab/transfer.py) lines 137, 185.

**Project meaning:** global scale $s$. Distinct implementation from M-022 (hand-rolled). NumPy median averages the two central values at even $n$.

---

## Scientific ambiguities and implementation/documentation discrepancies

1. **$L_{\mathrm{coll}}$ is computed and exported but not applied.** Production $\hat y$ uses only $L_{\mathrm{instr}}$ (M-021). The name “two_log_ratio” can be read as a product $e^{L_{\mathrm{coll}}+L_{\mathrm{instr}}}$; that product is **not** implemented. At matched unclipped anchors the unused identity $y^{\mathrm{ORCH}}_{\mathrm{sib}}e^{L_{\mathrm{instr}}}=y^{\mathrm{IOWA}}_{\mathrm{tgt}}e^{L_{\mathrm{coll}}}$ holds because the predictor starts from the ORCH donor. It is not generalised to interpolated/clipped MIDI.

2. **“Media” is family-specific, but production surfaces now share one resolver (M-027).** String mean vs woodwind combination remain distinct. Preparatory $(I+O)/2$ (M-028) is still an offline diagnostic. Geometric means (M-046) are **summaries**, not the cell-wise Media. Historical violin Zenodo AVERAGE cells were **not** rewritten.

3. **YAML combination methods `fixed_weight` and `validation_optimised` are rejected** with `ConfigurationError` until specified. `cfg.transfer_weight` must equal $1-w_E$. `optimal_weight()` is still not a production optimisation workflow.

4. **Methodology vs `equal_weight`.** Calibration export says a genuine measurement is never replaced; `equal_weight` / `legacy_equal_weight` **do** replace it by a convex combination (M-026).

5. **Shared $L$ is not replication.** IOWA effect $= y_{\mathrm{arco}}e^{L}$ with $L$ from Orchidea (or McGill prediction). Averaging IOWA+ORCH does not add an independent experiment.

6. **Origin `measured` is a tag.** Copied anchors (`copy_anchors=True`), Iowa compiled cells, and family-transfer Iowa (M-024) can all be labelled in ways that affect M-007/M-027. Woodwind ORCH is a **deprecated alias** for a transfer estimate; string ORCH effects may be true recordings.

7. **F-061 is external.** Logs are of the stored metric. They are **not** dB SPL. Units are **not defined** in this repository.

8. **`pi95_*` are not computed.** Fields exist for paste/export only. QA text that says “carry PI95” does not generate intervals.

9. **Fill vs transfer interpolation still disagree by design.** M-016 holds edge $L$; M-032 PCHIP may extrapolate $y$ and now tags it `extrapolated_pchip`. Exterior linear hold is `extrapolated_hold`, not interpolated. The scientific edge policy was not changed merely to make the two operators match.

10. **MedAE $\neq$ median.** `sorted(abs_err)[n//2]` (M-035) vs averaged two-middle median (M-022).

11. **`combined_empirical_only` validation row is a duplicate of the Iowa model row** (`build_validation_tables` line 689).

12. **Catalogue `transposition` is unused** in MIDI math (M-004). Written-pitch pastes will not be auto-corrected.

13. **Two harmonic floors** (catalogue `harmonics_sounding_low` vs pipeline `HARM_FLOOR`) differ by about an octave.

14. **`CalibrationConfig.seed` is unused.** No RNG in the documented estimators.

15. **`ln`RMSE / shape metrics are diagnostics**, not posterior intervals. No confidence or prediction interval is implemented.

16. **Fill documentation now matches the capped-exterior behaviour.** Methodology and the user manual no longer claim Fill is interior-only.

17. **Batch `linear_ratio` / GUI family transfer** are real code paths but not the woodwind production default.

18. **Instrument-specific batch CLIs** (`run_ste_effects_batch_*.py`) duplicate `transfer_or_copy` + harmonic floors; they are legacy/offline siblings of `run_ste_from_preparatory`, not different $L$ laws.

19. **`equal_weight` still mixes $y_E$ and $y_T$.** That is supported and labelled `COMBINED_ESTIMATE`. It is no longer advertised as if every method were empirical-only.

---

## Coverage appendix

| File | Lines | Classification |
|------|------:|----------------|
| [ste_lab/transfer.py](../ste_lab/transfer.py) | 368 | Project mathematics documented (M-015–M-019, M-025, M-032); library L-001–L-003, L-010 |
| [ste_lab/media_policy.py](../ste_lab/media_policy.py) | 196 | Production Media resolver (M-027) |
| [ste_lab/calibration.py](../ste_lab/calibration.py) | 884 | Project mathematics documented (M-021–M-023, M-026, M-033–M-042, M-050) |
| [ste_lab/calibration_export.py](../ste_lab/calibration_export.py) | 492 | Export of already-defined quantities; methodology prose cited in ambiguities |
| [ste_lab/notes.py](../ste_lab/notes.py) | 150 | Project mathematics documented (M-001–M-003) |
| [ste_lab/catalog.py](../ste_lab/catalog.py) | 268 | Project mathematics documented (M-004, M-031 ranks) |
| [ste_lab/session.py](../ste_lab/session.py) | 198 | Project mathematics documented (M-005, M-008, M-031) |
| [ste_lab/empirical.py](../ste_lab/empirical.py) | 186 | Project mathematics documented (M-012, M-013) |
| [ste_lab/evidence.py](../ste_lab/evidence.py) | 304 | Project mathematics documented (M-049) |
| [ste_lab/qa.py](../ste_lab/qa.py) | 443 | Project mathematics documented (M-043–M-045); remaining flags are non-numeric policy |
| [ste_lab/excel_export.py](../ste_lab/excel_export.py) | 805 | Project mathematics documented (M-046, M-047); woodwind wiring of M-021/M-026 |
| [ste_lab/zenodo_export.py](../ste_lab/zenodo_export.py) | 1614 | Project mathematics documented (M-007–M-009, M-029, M-030, M-046, M-048); L-005–L-008 |
| [ste_lab/paste.py](../ste_lab/paste.py) | 283 | Project mathematics documented (M-014); remainder I/O |
| [ste_lab/pipeline.py](../ste_lab/pipeline.py) | 157 | Non-mathematical dispatch |
| [ste_lab/gui_app.py](../ste_lab/gui_app.py) | 882 | Calls M-018/M-027/M-032; help text (F-061 not brightness); no extra estimator |
| [ste_lab/prep_gui.py](../ste_lab/prep_gui.py) | 473 | Non-mathematical GUI |
| [ste_lab/media_format.py](../ste_lab/media_format.py) | 104 | Non-mathematical (colours) |
| [ste_lab/__init__.py](../ste_lab/__init__.py) | 12 | Version 1.3.3 / policy comment |
| [ste_lab/__main__.py](../ste_lab/__main__.py) | 4 | Non-mathematical |
| [build_ste_preparatory.py](../build_ste_preparatory.py) | 1218 | Project mathematics documented (M-006, M-020, M-021, M-024, M-028) |
| [build_viola_ste_preparatory.py](../build_viola_ste_preparatory.py) | 19 | Duplicate/legacy wrapper; `HARM_LO=72` documented under M-006 |
| [run_ste_from_preparatory.py](../run_ste_from_preparatory.py) | 662 | Project mathematics documented (M-006, M-011, M-018, M-027) |
| [run_ste_effects_batch.py](../run_ste_effects_batch.py) | 691 | Project mathematics documented (M-010, M-018) |
| [run_ste_effects_batch_cello.py](../run_ste_effects_batch_cello.py) | 326 | Duplicate/legacy; floors `OPEN_C2=36`, `HARM_LO=60` under M-006 |
| [run_ste_effects_batch_viola.py](../run_ste_effects_batch_viola.py) | 331 | Duplicate/legacy of batch + floors |
| [run_ste_effects_batch_double_bass.py](../run_ste_effects_batch_double_bass.py) | 328 | Duplicate/legacy; `HARM_LO=52` |
| [run_ste_effects_batch_violin4.py](../run_ste_effects_batch_violin4.py) | 127 | Loader variant of M-010 (`spectral mass` column key) |
| [launch_ste_lab.py](../launch_ste_lab.py) | 14 | Non-mathematical |
| [launch_ste_prep.py](../launch_ste_prep.py) | 14 | Non-mathematical |
| [test_calibration.py](../test_calibration.py) | 470 | Verification evidence only (no independent production law) |
| [test_export_consistency.py](../test_export_consistency.py) | 539 | Export / provenance / Fill / config regressions |
| [test_run_ste_from_preparatory.py](../test_run_ste_from_preparatory.py) | 652 | Verification evidence |
| [test_pipeline_regression.py](../test_pipeline_regression.py) | 588 | Verification evidence |
| [calibration.yaml](../calibration.yaml) | 31 | Defaults cited throughout |
| `STE_Lab_User_and_Technical_Manual.html` | 593 | Family Media, combination modes, Fill labels, Ratio_extrapol identity |
| Notebooks | — | None present |
| Generated spreadsheet **formulas** | — | Documented (M-030, L-005–L-008) |
| `ste_lab/Viola_Zenodo_collections_harmonics.xlsx` | binary | Data/lookup; no analytical formula invented |

No first-party `.py` file was left “not inspected.”

---

## Appendix B — Source-file hashes

SHA-256 of the working-tree files after the export-consistency documentation pass (branch `fix/ratio-export-consistency`, starting commit `42a4df1`). Line counts are physical lines including blanks.

| File | SHA-256 | Lines |
|------|---------|------:|
| `README.md` | `21139e10334bcfcb890b22838a48e9d66607f24059aeb895e699e632119797f6` | 61 |
| `build_ste_preparatory.py` | `2ae13ab3b3eff9e259c5de1d43f04554ea0aa6c20063b8d612b1b43081c472ff` | 1218 |
| `build_viola_ste_preparatory.py` | `5987f3e4838aee00f8fd8e76cb6b4479eda85f76aa4d8f48d1f879a4caa9da8f` | 19 |
| `calibration.yaml` | `8c5e20732b789544e7c380ff8c30150752aedf11c1945aa63a49407e3fae4525` | 31 |
| `launch_ste_lab.py` | `dc0eb964bba2e20be9295b215f12b72f58bd2f98184ad30bc280db8a21df3971` | 14 |
| `launch_ste_prep.py` | `ce04439011a9431d23466a4488b68b3ab7a872f650f5da62f928a620cd3a9a76` | 14 |
| `requirements.txt` | `ab8b440bf65ba9e0837e16cc9ccd41910fd20e19995b433c8dea9c8e82b1f51b` | 4 |
| `run_ste_effects_batch.py` | `9dff744f5702fdb4488a137bb21d3bdbfa7ea53b9319e2cadc730cb3ce571687` | 691 |
| `run_ste_effects_batch_cello.py` | `02cb0f4987c582b0d2d64041dfae1edbf3ce6c3ff21ccf240fae653c95f3f647` | 326 |
| `run_ste_effects_batch_double_bass.py` | `fc024836275964c81b7f00110f0d848710222aca4d77f6c0fab8f73c4ac466c4` | 328 |
| `run_ste_effects_batch_viola.py` | `28ad8f4428c3e81b6355fb69ce1126407dd0018d725db586c08f92427fa25086` | 331 |
| `run_ste_effects_batch_violin4.py` | `1bf16ba86888a7164ffb02c68bd1f1f37dbc6353bcaff03ea056c57bbc2e155a` | 127 |
| `run_ste_from_preparatory.py` | `fc99bdffdde59d5dc5bb5d19f13fdb03e3f8e86766a457291b8999a04cb78a95` | 662 |
| `STE_Lab_User_and_Technical_Manual.html` | `b72130afc5446a118ab230e2be22ac2c23c483574871ddda0615c9f0adeba055` | 593 |
| `docs/README.md` | `5846eb94595e85b2ef5363a175cfd19dd2acade6b647941c94d813ff93bdbcf2` | 8 |
| `ste_lab/__init__.py` | `7d4a716eb9f072b3e32b3bc88d274826f3ff361ec2697e0a18cbd294685404c7` | 12 |
| `ste_lab/__main__.py` | `3d57e50a898a83eac0cb1ddbf6c181c726bb6ae54d70e66b601c5f37dcde9a6f` | 4 |
| `ste_lab/calibration.py` | `00717a615b6f7ba533a2db71a5833c441af9dff03a813c63755db2f6fa79eff2` | 884 |
| `ste_lab/calibration_export.py` | `317df806ca3e3681c5fec3278a8a91887ef7c4e3f77aff11eeb8b82452b32098` | 492 |
| `ste_lab/catalog.py` | `be07ad61d85268df7d395a23f0218ca00b11632f259e19fb23e0b5d4a4d2d763` | 268 |
| `ste_lab/empirical.py` | `af477263eb39c29bc28504b205e7ab70afc25f04b3dfb5848a231330b7f241c5` | 186 |
| `ste_lab/evidence.py` | `4385894f91180205abb2848ceb31646aa095c27dbceed76cbddf0361fb6db10a` | 304 |
| `ste_lab/excel_export.py` | `ee6366252310fe1452d28487e04d1e882ab4d3febc197aec1cd078d271f22438` | 805 |
| `ste_lab/gui_app.py` | `f3cbb2f8e21a2ddda74daae1def44d4d0339f3e1b1480187aa4e0947b9b722e0` | 882 |
| `ste_lab/media_format.py` | `81dc71318fbc21ea82f4e557cc87e2fb9585acc14ce04ce875c5a1801d802233` | 104 |
| `ste_lab/media_policy.py` | `2d5c19f556b2f716e481e2818d4002b2b5909001774339c1dd38440c864a11e5` | 196 |
| `ste_lab/notes.py` | `36eac89a505901a86dae543d7932eb5ab97e0486545fb343fe7ecdddceb80e5a` | 150 |
| `ste_lab/paste.py` | `56cb8638fd3fd066453bfb2d7f3e47ad794fc203b3b41523b24b408ffc8f1294` | 283 |
| `ste_lab/pipeline.py` | `5fe99db4d06af0815a06176335639494201773da08e784bba55aaa109d2046be` | 157 |
| `ste_lab/prep_gui.py` | `a4609ef9481d082c0553d4dd5df93008fea3c6c1e7f52e9b4eaa5626a6d059bd` | 473 |
| `ste_lab/qa.py` | `36bf2f19c58ee92b540e702d2422461307f24653b916c75b1311e196f8d1362b` | 443 |
| `ste_lab/session.py` | `6f93b0fc887b69c05ab81d12889886c53e63e4cf74200e3c6b43c6175e0bea02` | 198 |
| `ste_lab/transfer.py` | `a0d6bdd826ef4bad44cda87cb3c72da82cd2465064f10a20a266443e0e929219` | 368 |
| `ste_lab/zenodo_export.py` | `702eebc2b932b7402167a10d6d53d0fd01036612bdcfef72f6b79abc978b8fa4` | 1614 |
| `test_calibration.py` | `f0e5154d4aa9666d50bc91ec989b3f413c69aaf548700301f9be79e9b38a319c` | 470 |
| `test_export_consistency.py` | `6ab6eae2511941514676be32a6346bcd198870678fc51936f04e6ec0679db970` | 539 |
| `test_pipeline_regression.py` | `31c01f19543b68f486f1d4885829297981a7a891e4f1d4ba616636c2c838600d` | 588 |
| `test_run_ste_from_preparatory.py` | `d0477d54abda3338edaf2d51bf122ef554ff763a58a2be5d60f22bf4fe91e832` | 652 |

---

## References actually consulted

- Local sources listed in the coverage appendix (read in full or traced from callers).
- [calibration.yaml](../calibration.yaml).
- [requirements.txt](../requirements.txt).
- Git: `git rev-parse`, `git status`, `git remote`, `git log -1` (no fetch).
- GitHub: `gh api repos/LuisMRaimundo/Ratio_extrapol`, `gh api .../commits/main`, `git ls-remote https://github.com/LuisMRaimundo/Ratio_extrapol.git`.
- SciPy 1.13.1 `PchipInterpolator` docstring via `inspect.getdoc` on the installed package.
- NumPy 2.2.6 behaviour of `numpy.log` (natural log), `numpy.interp` (edge hold), `numpy.median` / `numpy.polyfit` (from installed library, not re-derived).
- Project `ste_lab/__init__.py` version 1.3.3.
- `STE_Lab_User_and_Technical_Manual.html` (F-061 wording only).
- Dynamics_extrapol `docs/Dynamics_extrapol_math_formula.md` used **only** as a formatting reference; **no** Dynamics formulae were copied as if they applied here.

The correction pass ran `python -m unittest discover -s . -p "test_*.py"` (116 tests, OK) on synthetic fixtures and the existing suite. It did **not** rerun the research corpus, regenerate historical workbooks, or render StackEdit. Isolated arithmetic checks: $e^{-3}\approx 0.049787$, MIDI(C4)$=60=12(4+1)+0$, $e^{3}\approx 20.0855$. These engineering fixes do **not** establish physical or acoustic validity.
