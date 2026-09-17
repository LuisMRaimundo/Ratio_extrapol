# -*- coding: utf-8 -*-
"""STE Lab batch for D:\\CORDAS_3\\VIOLIN 4.

Reads compiled F-061 from sheet Spectral_Density_Metrics, column
spectral_mass / spectral mass. Arco Media from the Zenodo collections
workbook. Writes STE + Zenodo workbooks to Extrapoled_effects.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import run_ste_effects_batch as batch
from run_ste_effects_batch import DYN_MAP, parse_pitch

VIOLIN4 = Path(r"D:\CORDAS_3\VIOLIN 4")
batch.ARCO = VIOLIN4 / "VIOLIN_Zenodo_collections_Arco_normal.xlsx"
batch.OUT = VIOLIN4 / "Extrapoled_effects"


def _col_key(name) -> str:
    return str(name).strip().lower().replace(" ", "_")


def _extract_mass(df: pd.DataFrame) -> dict[int, tuple[str, float]]:
    if df is None or df.empty:
        return {}
    note_col = next((c for c in df.columns if _col_key(c) == "note"), None)
    mass_col = next((c for c in df.columns if _col_key(c) == "spectral_mass"), None)
    if note_col is None or mass_col is None:
        return {}
    out: dict[int, tuple[str, float]] = {}
    for _, row in df.iterrows():
        pitch = parse_pitch(row.get(note_col))
        try:
            value = float(row.get(mass_col))
        except (TypeError, ValueError):
            continue
        if pitch is None or value <= 0:
            continue
        out[pitch.midi] = (pitch.label, value)
    return out


def load_spectral_mass(path: Path) -> dict[int, tuple[str, float]]:
    xl = pd.ExcelFile(path)
    if "Spectral_Density_Metrics" not in xl.sheet_names:
        print(f"  SKIP no Spectral_Density_Metrics: {path}")
        return {}
    return _extract_mass(pd.read_excel(path, sheet_name="Spectral_Density_Metrics"))


def discover_research_files() -> list[dict]:
    specs = []
    warnings: list[str] = []
    warnings.extend(batch.audit_iowa_zenodo_workbook(batch.ARCO, media_sheet="Violin_Media"))
    trees = [
        (VIOLIN4 / "Orchidea_violin", "ORCH", None),
        (VIOLIN4 / "Philharmonia_violin", "PHIL", None),
        (VIOLIN4 / "McGill_violin", "MCGILL", None),
    ]
    for tree, collection, forced_tech in trees:
        warnings.extend(
            batch.audit_research_tree(tree, collection=collection, technique=forced_tech or "")
        )
        if not tree.exists():
            continue
        for path in tree.rglob("*compiled_density_metrics_research.xlsx"):
            if path.name.startswith("~$"):
                continue
            parts = [p.lower() for p in path.parts]
            path_l = str(path).lower()
            if "_sustains" in path_l and "_sustains_stable" not in path_l:
                continue
            tech = forced_tech
            if tech is None:
                if "con-sord" in parts or "muted" in parts:
                    tech = "con sordino"
                elif "sul-ponticello" in parts or "arco-sul-ponticello" in parts:
                    tech = "sul ponticello"
                elif "sul-tasto" in parts or "arco-sul-tasto" in parts:
                    tech = "sul tasto"
                elif "harmonics" in parts:
                    tech = "harmonics"
                elif "arco-normal" in parts or "non-vibrato" in parts or "ordinario" in parts:
                    tech = "ordinario"
                else:
                    continue
            dyn = None
            for folder, code in DYN_MAP.items():
                if folder in parts:
                    dyn = code
                    break
            if dyn is None and collection == "MCGILL":
                dyn = "mf"
            if dyn is None:
                continue
            specs.append(
                {
                    "path": path,
                    "collection": collection,
                    "technique": tech,
                    "dynamic": dyn,
                }
            )
    batch.log_discovered_specs(specs)
    discover_research_files.last_warnings = warnings
    return specs


batch._extract_mass = _extract_mass
batch.load_spectral_mass = load_spectral_mass
batch.discover_research_files = discover_research_files


if __name__ == "__main__":
    batch.OUT.mkdir(parents=True, exist_ok=True)
    print("ARCO", batch.ARCO, "exists", batch.ARCO.exists())
    print("OUT ", batch.OUT)
    batch.main()
