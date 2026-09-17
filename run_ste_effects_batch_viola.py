# -*- coding: utf-8 -*-
"""Batch STE Lab run: IOWA + ORCHIDEA for every viola effect we can ground in data."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from run_ste_effects_batch import (
    CORE_DYNS,
    DYN_MAP,
    anchors_from_pair,
    audit_iowa_zenodo_workbook,
    audit_research_tree,
    build_layer,
    load_arco_media,
    load_spectral_mass,
    log_discovered_specs,
    transfer_or_copy,
)
from ste_lab.evidence import stamp_evidence
from ste_lab.excel_export import export_workbook
from ste_lab.notes import parse_pitch
from ste_lab.qa import audit_project
from ste_lab.session import Project
from ste_lab.media_policy import production_media_of
from ste_lab.zenodo_export import export_zenodo_workbook

VIOLA4 = Path(r"D:\CORDAS_3\VIOLA 4")
ARCO = VIOLA4 / "VIOLA_Zenodo_collections_Arco_normal.xlsx"
TREE = VIOLA4
OUT = VIOLA4 / "Viola_Extrapoled_effects"


def discover_research_files() -> list[dict]:
    specs = []
    warnings: list[str] = []
    warnings.extend(audit_iowa_zenodo_workbook(ARCO, media_sheet="VIOLA_Media"))
    trees = [
        (TREE / "ORCH_VLA" / "sul-ponticello", "ORCH", "sul ponticello"),
        (TREE / "ORCH_VLA" / "harmonics", "ORCH", "harmonics"),
        (TREE / "ORCH_VLA" / "ordinario", "ORCH", "ordinario"),
        (TREE / "Philharmonia_viola", "PHIL", None),
        (TREE / "McGill_viola", "MCGILL", None),
    ]
    for tree, collection, forced_tech in trees:
        warnings.extend(
            audit_research_tree(tree, collection=collection, technique=forced_tech or "")
        )
        if not tree.exists():
            continue
        for path in tree.rglob("*compiled_density_metrics_research.xlsx"):
            if path.name.startswith("~$"):
                continue
            parts = [p.lower() for p in path.parts]
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
                elif "arco-normal" in parts or "non-vibrato" in parts:
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
            if "_sustains\\" in str(path).lower() and "_sustains_stable" not in str(path).lower():
                continue
            specs.append(
                {
                    "path": path,
                    "collection": collection,
                    "technique": tech,
                    "dynamic": dyn,
                }
            )
    log_discovered_specs(specs)
    discover_research_files.last_warnings = warnings
    return specs


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    print("Loading viola arco-normal Media…")
    print("  ARCO", ARCO)
    print("  OUT ", OUT)
    arco = load_arco_media(ARCO)
    print(
        "  IOWA", {d: len(arco["IOWA"][d]) for d in CORE_DYNS},
        "ORCH", {d: len(arco["ORCH"][d]) for d in CORE_DYNS},
    )

    print("Discovering research workbooks…")
    specs = discover_research_files()
    measured: dict[tuple[str, str, str], dict] = {}
    for spec in specs:
        mass = load_spectral_mass(spec["path"])
        key = (spec["collection"], spec["technique"], spec["dynamic"])
        print(f"  {spec['collection']:6} {spec['technique']:16} {spec['dynamic']:3} n={len(mass)}  {spec['path'].name}")
        if not mass:
            continue
        if key not in measured or len(mass) > len(measured[key]["curve"]):
            measured[key] = {"curve": mass, "path": spec["path"]}

    # No viola sul tasto: Philharmonia_viola has no tasto tree, and STE must not invent one.
    effects = ["sul ponticello", "harmonics", "con sordino"]
    written = []

    for effect in effects:
        print(f"\n=== {effect} ===")
        project = Project(
            title=f"Viola {effect} — STE Lab IOWA + ORCHIDEA",
            operator="STE Lab batch",
            notes=(
                "Measured spectral_mass (F-061) from compiled_density_metrics_research.xlsx. "
                "Principal evidence = Empirical_ORCH (measured Orchidea only). "
                "IOWA has no effect recordings: IOWA = IOWA_ordinario × exp(L), L from Orchidea "
                "effect/ordinario pairs when available — shared L, not replication. "
                "Philharmonia/McGill are context layers, not Media. "
                "Con sordino has no Orchidea viola recordings: L from McGill muted is prediction."
            ),
        )

        orch_meas = {d: measured.get(("ORCH", effect, d), {}).get("curve", {}) for d in CORE_DYNS + ("p", "mp", "f")}
        phil_meas = {d: measured.get(("PHIL", effect, d), {}).get("curve", {}) for d in list(DYN_MAP.values())}
        mcgill_meas = {d: measured.get(("MCGILL", effect, d), {}).get("curve", {}) for d in CORE_DYNS + ("p",)}

        teacher = {}
        for d in CORE_DYNS:
            teacher[d] = orch_meas.get(d) or {}
        if not any(teacher.values()):
            if any(mcgill_meas.values()):
                print("  No Orchidea measurements; using McGill as L teacher (context-grade).")
                for d, curve in mcgill_meas.items():
                    if curve:
                        teacher[d] = curve
            elif any(phil_meas.values()):
                print("  No Orchidea measurements; using Philharmonia as L teacher (context-grade).")
                for d, curve in phil_meas.items():
                    if curve:
                        teacher[d] = curve
            else:
                print("  No teacher curve — skip")
                continue

        donor_dyn = next((d for d in ("mf", "pp", "ff", "p") if teacher.get(d)), "mf")
        donor_curve = teacher[donor_dyn]
        if effect == "con sordino" and mcgill_meas.get("mf"):
            mcgill_arco = measured.get(("MCGILL", "ordinario", "mf"), {}).get("curve", {})
            donor_arco_curve = {m: v[1] if isinstance(v, tuple) else v for m, v in mcgill_arco.items()}
            donor_source_label = "MCGILL ordinario mf"
        else:
            donor_arco_curve = arco["ORCH"].get(donor_dyn) or arco["ORCH"]["mf"]
            donor_source_label = f"ORCH ordinario {donor_dyn}"
        donor_anchors = anchors_from_pair(donor_arco_curve, donor_curve)
        print(f"  L donor {donor_dyn} vs {donor_source_label}: {len(donor_anchors)} anchors")

        for dyn in CORE_DYNS:
            iowa_arco_layer = build_layer(
                "viola", "IOWA", "ordinario", dyn, arco["IOWA"][dyn], "measured",
                "VIOLA_Media arco-normal", str(ARCO.name),
            )
            orch_arco_layer = build_layer(
                "viola", "ORCH", "ordinario", dyn, arco["ORCH"][dyn], "measured",
                "VIOLA_Media arco-normal", str(ARCO.name),
            )

            used_donor = True
            if teacher.get(dyn) and orch_meas.get(dyn):
                local_arco = arco["ORCH"].get(dyn) or arco["ORCH"]["mf"]
                anchors = anchors_from_pair(local_arco, teacher[dyn])
                used_donor = False
            else:
                anchors = donor_anchors
            if len(anchors) < 3:
                anchors = donor_anchors
                used_donor = True
            inherited_from = donor_dyn if used_donor else None
            prediction = effect == "con sordino" and not any(orch_meas.values())
            has_orch_meas = bool(orch_meas.get(dyn))
            min_midi = 72 if effect == "harmonics" else None

            if has_orch_meas:
                orch_eff = build_layer(
                    "viola",
                    "ORCH",
                    effect,
                    dyn,
                    orch_meas[dyn],
                    "measured",
                    "Orchidea compiled spectral_mass F-061",
                    str(measured[("ORCH", effect, dyn)]["path"].name),
                )
                if min_midi is not None:
                    for midi in [m for m in orch_eff.cells if m < min_midi]:
                        orch_eff.cells.pop(midi, None)
                transferred = transfer_or_copy(
                    orch_arco_layer,
                    anchors,
                    effect,
                    "ORCH",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if transferred:
                    for midi, cell in transferred.cells.items():
                        if midi not in orch_eff.cells:
                            orch_eff.place(
                                parse_pitch(f"midi{midi}"),
                                cell.value,
                                "modelled_ORCHIDEA_anchored",
                                "filled from ordinario × exp(L)",
                                overwrite="overwrite",
                            )
                stamp_evidence(orch_eff, inherited_from=inherited_from, prediction=prediction)
                project.add_layer(orch_eff)
            else:
                transferred = transfer_or_copy(
                    orch_arco_layer,
                    anchors,
                    effect,
                    "ORCH",
                    "modelled_ORCHIDEA_anchored",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if transferred:
                    stamp_evidence(transferred, inherited_from=inherited_from or donor_dyn, prediction=prediction)
                    project.add_layer(transferred)

            iowa_eff = transfer_or_copy(
                iowa_arco_layer,
                anchors,
                effect,
                "IOWA",
                "modelled_IOWA_anchored",
                copy_anchors=False,
                min_midi=min_midi,
            )
            if iowa_eff:
                if not orch_meas.get(dyn):
                    for cell in iowa_eff.cells.values():
                        if cell.midi > 88:
                            cell.reporting_status = "above_measured_ceiling"
                stamp_evidence(iowa_eff, inherited_from=inherited_from, prediction=prediction)
                project.add_layer(iowa_eff)

        for dyn, curve in phil_meas.items():
            if not curve:
                continue
            src = measured.get(("PHIL", effect, dyn), {}).get("path")
            project.add_layer(
                build_layer(
                    "viola", "PHIL", effect, dyn, curve, "measured",
                    "Philharmonia context — not pooled into Media",
                    src.name if src else "",
                )
            )
        for dyn, curve in mcgill_meas.items():
            if not curve:
                continue
            src = measured.get(("MCGILL", effect, dyn), {}).get("path")
            project.add_layer(
                build_layer(
                    "viola", "MCGILL", effect, dyn, curve, "measured",
                    "McGill context — not pooled into Media",
                    src.name if src else "",
                )
            )

        media_layers = []
        for dyn in CORE_DYNS:
            peers = [
                lg
                for lg in project.layers
                if lg.technique == effect
                and lg.dynamic == dyn
                and lg.collection in {"IOWA", "ORCH"}
                and lg.cells
            ]
            if len(peers) >= 1:
                media = production_media_of(
                    peers,
                    name=f"viola_Media_{effect}_{dyn}",
                    instrument="viola",
                    technique=effect,
                )
                project.add_layer(media)
                media_layers.append(media)

        flags = audit_project(project)
        slug = effect.replace(" ", "_")
        out = OUT / f"Viola_STE_{slug}_IOWA_ORCH.xlsx"
        export_workbook(project, out, flags, media_layers)
        written.append(out)
        print(f"  wrote {out.name}  layers={len(project.layers)}  flags={len(flags)}  size={out.stat().st_size}")

        zenodo = OUT / f"Viola_Zenodo_collections_{slug}.xlsx"
        export_zenodo_workbook(
            project,
            zenodo,
            technique=effect,
            arco=arco,
            notes=project.notes,
        )
        written.append(zenodo)
        print(f"  wrote {zenodo.name}  size={zenodo.stat().st_size}")

    print("\nDone:")
    for p in written:
        print(" ", p)


if __name__ == "__main__":
    main()
