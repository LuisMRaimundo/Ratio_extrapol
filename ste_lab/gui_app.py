"""STE Lab desktop GUI — data entry, chromatic snap, transfer, QA, Excel export."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser
from pathlib import Path

from . import __version__
from .catalog import (
    COLLECTIONS,
    DYNAMICS,
    FAMILIES,
    ORIGINS,
    PITCH_BASES,
    TECHNIQUES,
    instrument_choices,
    resolve_instrument,
)
from .excel_export import export_workbook, import_workbook
from .media_policy import production_average_for_layer
from .notes import chromatic_range, midi_to_label, parse_pitch
from .paste import parse_block
from .qa import audit_project, audit_transfer
from .session import Project, layer_from_identity
from .transfer import InsufficientFillData, family_transfer, fill_missing, parse_anchor_block, technique_transfer


APP_TITLE = "STE Lab — Spectral Technique Extrapolation"
MANUAL_NAME = "STE_Lab_User_and_Technical_Manual.html"


def manual_path() -> Path:
    return Path(__file__).resolve().parents[1] / MANUAL_NAME


def _open_local_html(path: Path) -> None:
    """Open a local HTML file. file:// URIs with accents make Edge search instead."""
    path = path.resolve()
    if os.name == "nt":
        os.startfile(str(path))
        return
    webbrowser.open(path.as_uri())


class STEApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1480x900")
        self.minsize(1180, 720)
        self.project = Project(title="Untitled STE session")
        self.current_layer_id: str | None = None
        self._build_style()
        self._build_ui()
        self._refresh_all()

    def _build_style(self) -> None:
        self.configure(bg="#eef2f6")
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#eef2f6")
        style.configure("TLabelframe", background="#eef2f6")
        style.configure("TLabelframe.Label", background="#eef2f6", font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", background="#1f4e79", foreground="white", font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel", background="#1f4e79", foreground="#d6e4f0", font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 9), padding=4)
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", font=("Consolas", 9), rowheight=22)
        style.configure("Treeview.Heading", font=("Segoe UI", 8, "bold"))

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg="#1f4e79")
        header.pack(fill="x")
        titles = tk.Frame(header, bg="#1f4e79")
        titles.pack(side="left", fill="x", expand=True)
        ttk.Label(titles, text=APP_TITLE, style="Header.TLabel").pack(anchor="w", padx=16, pady=(10, 0))
        ttk.Label(
            titles,
            text="Measured cells stay measured. Pasted blocks snap to concert pitch. Risky methods are flagged before export.",
            style="Sub.TLabel",
        ).pack(anchor="w", padx=16, pady=(0, 10))
        help_btn = tk.Button(
            header,
            text="Help",
            command=self._open_help,
            font=("Segoe UI", 10, "bold"),
            bg="#d6e4f0",
            fg="#1f4e79",
            activebackground="#ffffff",
            activeforeground="#1f4e79",
            relief="flat",
            padx=16,
            pady=6,
            cursor="hand2",
        )
        help_btn.pack(side="right", padx=16, pady=14)
        self.bind("<F1>", lambda e: self._open_help())
        self.bind("<Help>", lambda e: self._open_help())

        meta = ttk.Frame(self)
        meta.pack(fill="x", padx=10, pady=6)
        ttk.Label(meta, text="Project").grid(row=0, column=0, sticky="w")
        self.var_title = tk.StringVar(value=self.project.title)
        ttk.Entry(meta, textvariable=self.var_title, width=42).grid(row=0, column=1, padx=6)
        ttk.Label(meta, text="Operator").grid(row=0, column=2, sticky="w")
        self.var_operator = tk.StringVar()
        ttk.Entry(meta, textvariable=self.var_operator, width=22).grid(row=0, column=3, padx=6)
        ttk.Label(meta, text="Session notes").grid(row=0, column=4, sticky="w")
        self.var_notes = tk.StringVar()
        ttk.Entry(meta, textvariable=self.var_notes, width=48).grid(row=0, column=5, padx=6, sticky="ew")
        meta.columnconfigure(5, weight=1)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=8, pady=4)

        left = ttk.Notebook(body, width=520)
        right = ttk.Panedwindow(body, orient="vertical")
        body.add(left, weight=2)
        body.add(right, weight=3)

        self.tab_identity = ttk.Frame(left)
        self.tab_entry = ttk.Frame(left)
        self.tab_transfer = ttk.Frame(left)
        self.tab_export = ttk.Frame(left)
        left.add(self.tab_identity, text="1. Identity & labels")
        left.add(self.tab_entry, text="2. Insert data")
        left.add(self.tab_transfer, text="3. Extrapolate")
        left.add(self.tab_export, text="4. Export")

        self._build_identity(self.tab_identity)
        self._build_entry(self.tab_entry)
        self._build_transfer(self.tab_transfer)
        self._build_export(self.tab_export)

        grid_wrap = ttk.Frame(right)
        warn_wrap = ttk.Frame(right)
        right.add(grid_wrap, weight=3)
        right.add(warn_wrap, weight=2)
        self._build_grid(grid_wrap)
        self._build_warnings(warn_wrap)

    def _combo(self, parent, values, width=22, default="") -> ttk.Combobox:
        box = ttk.Combobox(parent, values=values, width=width)
        if default:
            box.set(default)
        elif values:
            box.set(values[0])
        return box

    def _labeled(self, parent, row, col, text, widget) -> None:
        ttk.Label(parent, text=text).grid(row=row, column=col, sticky="w", padx=4, pady=2)
        widget.grid(row=row, column=col + 1, sticky="ew", padx=4, pady=2)

    def _build_identity(self, parent: ttk.Frame) -> None:
        src = ttk.LabelFrame(parent, text="Source layer (what you measured or already have)")
        src.pack(fill="x", padx=8, pady=8)
        tgt = ttk.LabelFrame(parent, text="Target layer (what you want to obtain)")
        tgt.pack(fill="x", padx=8, pady=8)
        extra = ttk.LabelFrame(parent, text="Free labels (printed on the Excel)")
        extra.pack(fill="both", expand=True, padx=8, pady=8)

        families = list(FAMILIES.values())
        self.src_family = self._combo(src, families, 24, FAMILIES["bowed_strings"])
        self.src_instr = self._combo(src, [s.display_name for s in instrument_choices("bowed_strings")], 24, "Viola")
        self.src_coll = self._combo(src, COLLECTIONS, 20, "IOWA")
        self.src_tech = self._combo(src, TECHNIQUES, 20, "ordinario")
        self.src_dyn = self._combo(src, DYNAMICS, 8, "mf")
        self.src_basis = self._combo(src, PITCH_BASES, 20, "sounding_concert")
        self._labeled(src, 0, 0, "Family", self.src_family)
        self._labeled(src, 0, 2, "Instrument", self.src_instr)
        self._labeled(src, 1, 0, "Collection", self.src_coll)
        self._labeled(src, 1, 2, "Technique", self.src_tech)
        self._labeled(src, 2, 0, "Dynamic", self.src_dyn)
        self._labeled(src, 2, 2, "Pitch basis", self.src_basis)
        self.src_family.bind("<<ComboboxSelected>>", lambda e: self._filter_instruments(self.src_family, self.src_instr))

        self.tgt_family = self._combo(tgt, families, 24, FAMILIES["bowed_strings"])
        self.tgt_instr = self._combo(tgt, [s.display_name for s in instrument_choices("bowed_strings")], 24, "Viola")
        self.tgt_coll = self._combo(tgt, COLLECTIONS, 20, "ORCH")
        self.tgt_tech = self._combo(tgt, TECHNIQUES, 20, "harmonics")
        self.tgt_dyn = self._combo(tgt, DYNAMICS, 8, "mf")
        self.tgt_basis = self._combo(tgt, PITCH_BASES, 20, "sounding_concert")
        self._labeled(tgt, 0, 0, "Family", self.tgt_family)
        self._labeled(tgt, 0, 2, "Instrument", self.tgt_instr)
        self._labeled(tgt, 1, 0, "Collection", self.tgt_coll)
        self._labeled(tgt, 1, 2, "Technique", self.tgt_tech)
        self._labeled(tgt, 2, 0, "Dynamic", self.tgt_dyn)
        self._labeled(tgt, 2, 2, "Pitch basis", self.tgt_basis)
        self.tgt_family.bind("<<ComboboxSelected>>", lambda e: self._filter_instruments(self.tgt_family, self.tgt_instr))

        ttk.Label(extra, text="Additional labels (key = value, one per line)").pack(anchor="w", padx=6)
        self.txt_labels = tk.Text(extra, height=5, font=("Consolas", 9))
        self.txt_labels.pack(fill="both", expand=True, padx=6, pady=4)
        self.txt_labels.insert("1.0", "sample_type = sustains\nmetric = Combined Density Metric\n")

        btns = ttk.Frame(parent)
        btns.pack(fill="x", padx=8, pady=6)
        ttk.Button(btns, text="New source layer", command=self._new_source_layer).pack(side="left", padx=3)
        ttk.Button(btns, text="New target layer (empty)", command=self._new_target_layer).pack(side="left", padx=3)
        ttk.Button(btns, text="Apply labels to selected layer", command=self._apply_labels).pack(side="left", padx=3)

    def _build_entry(self, parent: ttk.Frame) -> None:
        one = ttk.LabelFrame(parent, text="Individual numerical cell")
        one.pack(fill="x", padx=8, pady=8)
        self.ent_note = self._combo(one, [p.label for p in chromatic_range(36, 108)], 10, "C4")
        self.var_value = tk.StringVar()
        self.ent_origin = self._combo(one, ORIGINS, 22, "measured")
        self.var_comment = tk.StringVar()
        self._labeled(one, 0, 0, "Note", self.ent_note)
        ttk.Label(one, text="CDM / value").grid(row=0, column=2, sticky="w")
        ttk.Entry(one, textvariable=self.var_value, width=12).grid(row=0, column=3, padx=4)
        self._labeled(one, 1, 0, "Origin", self.ent_origin)
        ttk.Label(one, text="Comment").grid(row=1, column=2, sticky="w")
        ttk.Entry(one, textvariable=self.var_comment, width=28).grid(row=1, column=3, padx=4, sticky="ew")
        ttk.Button(one, text="Place on chromatic grid", style="Accent.TButton", command=self._place_one).grid(
            row=2, column=0, columnspan=4, pady=6
        )

        block = ttk.LabelFrame(
            parent,
            text="Block paste — Excel, two columns, or a bare number list. Notes are sorted and slotted automatically.",
        )
        block.pack(fill="both", expand=True, padx=8, pady=8)
        self.txt_paste = tk.Text(block, height=14, font=("Consolas", 9))
        self.txt_paste.pack(fill="both", expand=True, padx=6, pady=4)
        self.txt_paste.insert(
            "1.0",
            "C5\t16.876\tmeasured\n"
            "Eb5\t17.080\tmeasured\n"
            "D5\t15.659\tmeasured\n"
            "Db5\t15.847\tmeasured\n",
        )
        row = ttk.Frame(block)
        row.pack(fill="x", padx=6, pady=4)
        ttk.Label(row, text="Start note if the paste is only numbers").pack(side="left")
        self.ent_start = self._combo(row, [p.label for p in chromatic_range(24, 108)], 8, "C4")
        self.ent_start.pack(side="left", padx=6)
        ttk.Button(row, text="Parse clipboard", command=self._paste_clipboard).pack(side="left", padx=3)
        ttk.Button(row, text="Parse box and place", style="Accent.TButton", command=self._parse_and_place).pack(
            side="left", padx=3
        )
        ttk.Button(row, text="Clear", command=self._clear_paste).pack(side="left", padx=3)
        self.var_overwrite = tk.StringVar(value="prefer_measured")
        ttk.Label(row, text="If slot occupied").pack(side="left", padx=(12, 2))
        ttk.Combobox(
            row,
            textvariable=self.var_overwrite,
            values=["prefer_measured", "overwrite", "skip"],
            width=16,
        ).pack(side="left")

    def _build_transfer(self, parent: ttk.Frame) -> None:
        fill = ttk.LabelFrame(parent, text="Fill missing notes on the selected layer")
        fill.pack(fill="x", padx=8, pady=8)
        self.cmb_fill = self._combo(fill, ["pchip", "linear", "polynomial"], 14, "pchip")
        self.var_extrap = tk.StringVar(value="12")
        self.var_degree = tk.StringVar(value="3")
        self._labeled(fill, 0, 0, "Method", self.cmb_fill)
        ttk.Label(fill, text="Max extrapolation (semitones)").grid(row=0, column=2, sticky="w")
        ttk.Entry(fill, textvariable=self.var_extrap, width=6).grid(row=0, column=3)
        ttk.Label(fill, text="Polynomial degree").grid(row=1, column=0, sticky="w", padx=4)
        ttk.Entry(fill, textvariable=self.var_degree, width=6).grid(row=1, column=1, sticky="w")
        ttk.Label(
            fill,
            text="PCHIP needs ≥3 known notes. Linear is valid for two notes. Exterior tags are extrapolated_pchip / extrapolated_hold, not ridge.",
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=4)
        ttk.Button(fill, text="Fill selected layer", command=self._run_fill).grid(row=3, column=2, columnspan=2, pady=4)

        tech = ttk.LabelFrame(
            parent,
            text="Technique transfer — ordinario → con sordino / sul tasto / harmonics / au talon / sul ponticello",
        )
        tech.pack(fill="x", padx=8, pady=8)
        self.cmb_tech_method = self._combo(tech, ["log_ratio", "linear_ratio"], 16, "log_ratio")
        self._labeled(tech, 0, 0, "Method", self.cmb_tech_method)
        ttk.Label(tech, text="Uses the Target technique box. log_ratio = source × exp(L).").grid(
            row=0, column=2, columnspan=2, sticky="w"
        )
        ttk.Label(tech, text="Anchors (note  sourceCDM  targetCDM), one per line").grid(
            row=1, column=0, columnspan=4, sticky="w", padx=4
        )
        self.txt_anchors = tk.Text(tech, height=5, font=("Consolas", 9))
        self.txt_anchors.grid(row=2, column=0, columnspan=4, sticky="ew", padx=6, pady=4)
        self.txt_anchors.insert("1.0", "C5  16.88  12.40\nG5  14.20  10.10\nC6  14.31  9.80\n")
        ttk.Button(tech, text="Run technique transfer → new layer", command=self._run_technique).grid(
            row=3, column=0, columnspan=4, pady=4
        )

        fam = ttk.LabelFrame(parent, text="Family-register transfer — e.g. clarinet → bass clarinet")
        fam.pack(fill="x", padx=8, pady=8)
        self.cmb_fam_mode = self._combo(fam, ["concert_pitch", "register_relative"], 20, "concert_pitch")
        self._labeled(fam, 0, 0, "Alignment", self.cmb_fam_mode)
        ttk.Label(fam, text="Uses the Target instrument. Same-family only.").grid(row=0, column=2, sticky="w")
        ttk.Button(fam, text="Run family transfer → new layer", command=self._run_family).grid(
            row=1, column=0, columnspan=3, pady=6
        )

        med = ttk.LabelFrame(
            parent,
            text="Completed-grid Media (not independent replication of L — Empirical_ORCH is principal evidence)",
        )
        med.pack(fill="x", padx=8, pady=8)
        ttk.Button(
            med,
            text="Production Media of IOWA/ORCH peers (string mean; woodwind combination policy)",
            command=self._run_media,
        ).pack(padx=6, pady=6)

    def _build_export(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(parent, text="Workbook (same organisation as the viola harmonics Excel)")
        box.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(
            box,
            text="Sheets: Empirical_ORCH (principal evidence), Evidence_Map, collection layers,\n"
            "completed-grid Media, Media_Uncertainty, Summary_Measured_Range, AcousticTable, QA_Flags.\n"
            "Media is not a second experiment: IOWA and ORCH share L. Tasto / inherited pp–ff stay off Empirical_ORCH. Phil/McGill are context.",
            justify="left",
        ).pack(anchor="w", padx=8, pady=8)
        ttk.Button(box, text="Import existing STE / viola-style Excel…", command=self._import_xlsx).pack(
            anchor="w", padx=8, pady=4
        )
        ttk.Button(box, text="Generate organised Excel…", style="Accent.TButton", command=self._export_xlsx).pack(
            anchor="w", padx=8, pady=8
        )
        ttk.Button(box, text="Open user + technical manual…", command=self._open_manual_file).pack(
            anchor="w", padx=8, pady=4
        )
        ttk.Label(box, text=f"STE Lab v{__version__}  ·  F1 = Help").pack(anchor="w", padx=8, pady=4)

    def _build_grid(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=4, pady=4)
        ttk.Label(top, text="Layers").pack(side="left")
        self.lst_layers = tk.Listbox(top, height=5, exportselection=False, font=("Segoe UI", 9), width=42)
        self.lst_layers.pack(side="left", fill="x", expand=True, padx=6)
        self.lst_layers.bind("<<ListboxSelect>>", lambda e: self._on_select_layer())
        side = ttk.Frame(top)
        side.pack(side="left")
        ttk.Button(side, text="Delete layer", command=self._delete_layer).pack(fill="x", pady=2)
        ttk.Button(side, text="Delete selected note", command=self._delete_note).pack(fill="x", pady=2)

        cols = ("note", "midi", "value", "origin", "status", "pi95", "flags", "comment")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headings = {
            "note": "Note",
            "midi": "MIDI",
            "value": "CDM / value",
            "origin": "Origin",
            "status": "Reporting",
            "pi95": "PI95",
            "flags": "Flags",
            "comment": "Comment",
        }
        widths = {
            "note": 60,
            "midi": 48,
            "value": 100,
            "origin": 180,
            "status": 140,
            "pi95": 110,
            "flags": 120,
            "comment": 180,
        }
        for col in cols:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")
        self.tree.tag_configure("measured", background="#c6efce")
        self.tree.tag_configure("generated", background="#bdd7ee")
        self.tree.tag_configure("interpolated", background="#c5d9f1")
        self.tree.tag_configure("technique_transfer", background="#e2d5f1")
        self.tree.tag_configure("family_transfer", background="#d9ead3")
        self.tree.tag_configure("collection_anchored", background="#fff2cc")
        self.tree.tag_configure("modelled", background="#fce4d6")
        self.tree.tag_configure("modelled_IOWA_anchored", background="#fce4d6")
        self.tree.tag_configure("modelled_ORCHIDEA_anchored", background="#fce4d6")
        self.tree.tag_configure("modelled_on_extrapolated_anchor", background="#f4b183")
        self.tree.tag_configure("extrapolated_ridge", background="#f8cbad")
        self.tree.tag_configure("extrapolated_polynomial", background="#f4b183")
        self.tree.tag_configure("warn", background="#ffe599")
        self.tree.tag_configure("bad", background="#ff8a80")
        scroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        scroll.pack(side="right", fill="y", pady=4)

    def _build_warnings(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Quality desk — approaches and notes that can scramble the result").pack(
            anchor="w", padx=6, pady=(6, 2)
        )
        cols = ("sev", "code", "note", "approach", "message", "advice")
        self.warn = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        for col, title, w in (
            ("sev", "Sev.", 70),
            ("code", "Code", 160),
            ("note", "Note", 60),
            ("approach", "Approach", 160),
            ("message", "Why it is unsafe", 360),
            ("advice", "What to do instead", 360),
        ):
            self.warn.heading(col, text=title)
            self.warn.column(col, width=w)
        self.warn.tag_configure("critical", background="#ff6b6b")
        self.warn.tag_configure("high", background="#f4b183")
        self.warn.tag_configure("medium", background="#ffe599")
        self.warn.tag_configure("low", background="#d9ead3")
        y = ttk.Scrollbar(parent, orient="vertical", command=self.warn.yview)
        self.warn.configure(yscrollcommand=y.set)
        self.warn.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        y.pack(side="right", fill="y", pady=4)

    def _family_id(self, display: str) -> str:
        for key, label in FAMILIES.items():
            if label == display:
                return key
        return "other"

    def _filter_instruments(self, family_box: ttk.Combobox, instr_box: ttk.Combobox) -> None:
        fid = self._family_id(family_box.get())
        names = [s.display_name for s in instrument_choices(fid)]
        instr_box["values"] = names
        if names:
            instr_box.set(names[0])

    def _id_of(self, display: str) -> str:
        spec = resolve_instrument(display)
        return spec.instrument_id if spec else display.lower().replace(" ", "_")

    def _extra_labels(self) -> dict[str, str]:
        out = {}
        for line in self.txt_labels.get("1.0", "end").splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                out[k.strip()] = v.strip()
        return out

    def _sync_project_meta(self) -> None:
        self.project.title = self.var_title.get().strip() or "Untitled STE session"
        self.project.operator = self.var_operator.get().strip()
        self.project.notes = self.var_notes.get().strip()

    def _current_layer(self):
        if not self.current_layer_id:
            return None
        return self.project.get_layer(self.current_layer_id)

    def _new_source_layer(self) -> None:
        layer = layer_from_identity(
            self._id_of(self.src_instr.get()),
            self.src_coll.get(),
            self.src_tech.get(),
            self.src_dyn.get(),
            self._family_id(self.src_family.get()),
            self.src_basis.get(),
            extra_labels=self._extra_labels(),
        )
        self.project.add_layer(layer)
        self.current_layer_id = layer.layer_id
        self._refresh_all()

    def _new_target_layer(self) -> None:
        layer = layer_from_identity(
            self._id_of(self.tgt_instr.get()),
            self.tgt_coll.get(),
            self.tgt_tech.get(),
            self.tgt_dyn.get(),
            self._family_id(self.tgt_family.get()),
            self.tgt_basis.get(),
            extra_labels=self._extra_labels(),
        )
        self.project.add_layer(layer)
        self.current_layer_id = layer.layer_id
        self._refresh_all()

    def _apply_labels(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        layer.instrument = self._id_of(self.src_instr.get())
        layer.family = self._family_id(self.src_family.get())
        layer.collection = self.src_coll.get()
        layer.technique = self.src_tech.get()
        layer.dynamic = self.src_dyn.get()
        layer.pitch_basis = self.src_basis.get()
        layer.labels.update(self._extra_labels())
        layer.name = f"{layer.instrument}_{layer.collection}_{layer.technique}_{layer.dynamic}"
        self.project.log.append(f"Relabelled {layer.display_name()}")
        self._refresh_all()

    def _require_layer(self):
        layer = self._current_layer()
        if layer is None:
            messagebox.showinfo("STE Lab", "Create or select a layer first (tab 1).")
        return layer

    def _place_one(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        pitch = parse_pitch(self.ent_note.get())
        if not pitch:
            messagebox.showerror("Note", f"Cannot parse note '{self.ent_note.get()}'.")
            return
        try:
            value = float(self.var_value.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Value", "Enter a numerical CDM / metric.")
            return
        status = layer.place(pitch, value, self.ent_origin.get(), self.var_comment.get(), self.var_overwrite.get())
        self.project.log.append(f"{status} {pitch.label}={value} on {layer.display_name()}")
        self._refresh_all()

    def _clear_paste(self) -> None:
        self.txt_paste.delete("1.0", "end")
        self.var_value.set("")
        self.var_comment.set("")

    def _paste_clipboard(self) -> None:
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Clipboard", "Clipboard is empty.")
            return
        self.txt_paste.delete("1.0", "end")
        self.txt_paste.insert("1.0", text)
        self._parse_and_place()

    def _parse_and_place(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        report = parse_block(self.txt_paste.get("1.0", "end"), self.ent_start.get())
        if report.start_note_needed:
            messagebox.showwarning("Start note needed", "\n".join(report.messages))
            return
        if not report.placed_pairs:
            messagebox.showwarning("Parse", "\n".join(report.messages) or "No note/value pairs found.")
            return
        counts = {"placed": 0, "overwritten": 0, "skipped": 0, "merged_kept": 0}
        for row in report.placed_pairs:
            origin = row.origin or self.ent_origin.get() or "manual"
            status = layer.place(
                row.pitch,
                row.value,
                origin,
                row.comment,
                self.var_overwrite.get(),
                pi95_low=row.pi95_low,
                pi95_high=row.pi95_high,
                reporting_status=row.reporting_status,
                anchor_source=row.anchor_source,
            )
            counts[status] = counts.get(status, 0) + 1
            if row.dynamic and row.dynamic.lower() != layer.dynamic.lower():
                layer.cells[row.pitch.midi].flags.append(f"pasted_dynamic={row.dynamic}")
        layer.range_low = min(layer.range_low, min(layer.cells))
        layer.range_high = max(layer.range_high, max(layer.cells))
        self.project.log.append(f"Block paste on {layer.display_name()}: {counts}")
        self._refresh_all()
        messagebox.showinfo("Placed", "\n".join(report.messages + [f"Grid update: {counts}"]))

    def _run_fill(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        try:
            cap = int(self.var_extrap.get())
            degree = int(self.var_degree.get())
        except ValueError:
            messagebox.showerror("Parameters", "Extrapolation cap and degree must be integers.")
            return
        pre = audit_transfer(layer, layer.instrument, layer.technique, self.cmb_fill.get(), len(layer.measured_midis()), cap, degree)
        if any(f.severity == "critical" for f in pre):
            if not messagebox.askyesno("Blocked approach", self._flag_text(pre) + "\n\nRun anyway?"):
                return
        try:
            result = fill_missing(layer, self.cmb_fill.get(), cap, degree)
        except InsufficientFillData as exc:
            messagebox.showerror(
                "Fill",
                f"{exc}\nThe selected layer was left unchanged. Choose linear for two notes, or add another measured value.",
            )
            return
        self.project.log.extend(result.messages)
        self._refresh_all()
        messagebox.showinfo("Fill", "\n".join(result.messages))

    def _run_technique(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        anchors = parse_anchor_block(self.txt_anchors.get("1.0", "end"))
        target_tech = self.tgt_tech.get()
        pre = audit_transfer(
            layer,
            layer.instrument,
            target_tech,
            self.cmb_tech_method.get(),
            len(anchors),
            int(self.var_extrap.get() or 12),
            int(self.var_degree.get() or 3),
        )
        if any(f.severity == "critical" for f in pre):
            if not messagebox.askyesno("Inadvisable technique transfer", self._flag_text(pre) + "\n\nContinue?"):
                return
        result = technique_transfer(layer, anchors, target_tech, self.cmb_tech_method.get(), self.tgt_coll.get())
        self.project.add_layer(result.layer)
        self.current_layer_id = result.layer.layer_id
        self.project.log.extend(result.messages)
        self._refresh_all()
        messagebox.showinfo("Technique transfer", "\n".join(result.messages + self._flag_text(pre).splitlines()[:6]))

    def _run_family(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        target = self._id_of(self.tgt_instr.get())
        anchors = parse_anchor_block(self.txt_anchors.get("1.0", "end"))
        pre = audit_transfer(
            layer,
            target,
            layer.technique,
            "family_" + self.cmb_fam_mode.get().split("_")[0],
            len(anchors),
            int(self.var_extrap.get() or 12),
            3,
        )
        if any(f.severity == "critical" for f in pre):
            messagebox.showerror("Cross-family transfer refused", self._flag_text(pre))
            return
        if pre and not messagebox.askyesno("Family transfer warnings", self._flag_text(pre) + "\n\nContinue?"):
            return
        result = family_transfer(layer, target, self.cmb_fam_mode.get(), anchors)
        self.project.add_layer(result.layer)
        self.current_layer_id = result.layer.layer_id
        self.project.log.extend(result.messages)
        self._refresh_all()
        messagebox.showinfo("Family transfer", "\n".join(result.messages))

    def _run_media(self) -> None:
        layer = self._require_layer()
        if not layer:
            return
        try:
            media = production_average_for_layer(self.project.layers, layer)
        except ValueError as exc:
            messagebox.showinfo("Media", str(exc))
            return
        self.project.add_layer(media)
        self.current_layer_id = media.layer_id
        self.project.log.append(f"Production Media → {media.display_name()}")
        self._refresh_all()

    def _export_xlsx(self) -> None:
        self._sync_project_meta()
        if not self.project.layers:
            messagebox.showinfo("Export", "Nothing to export yet.")
            return
        path = filedialog.asksaveasfilename(
            title="Save STE workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{self.project.layers[0].instrument}_{self.project.layers[0].technique}_STE.xlsx",
        )
        if not path:
            return
        media_layers = [lg for lg in self.project.layers if lg.collection.startswith("+") or "Media" in lg.name]
        flags = audit_project(self.project)
        export_workbook(self.project, Path(path), flags, media_layers)
        self.project.log.append(f"Exported {path}")
        messagebox.showinfo("Exported", f"Workbook written:\n{path}\nQA flags: {len(flags)}")

    def _import_xlsx(self) -> None:
        path = filedialog.askopenfilename(title="Open STE / viola-style Excel", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        n = import_workbook(Path(path), self.project)
        if n and not self.current_layer_id:
            self.current_layer_id = self.project.layers[0].layer_id
        self._refresh_all()
        messagebox.showinfo("Import", f"Loaded {n} collection layer(s) from\n{path}")

    def _delete_layer(self) -> None:
        layer = self._current_layer()
        if not layer:
            return
        if not messagebox.askyesno("Delete", f"Delete layer {layer.display_name()}?"):
            return
        self.project.remove_layer(layer.layer_id)
        self.current_layer_id = self.project.layers[0].layer_id if self.project.layers else None
        self._refresh_all()

    def _delete_note(self) -> None:
        layer = self._current_layer()
        if not layer:
            return
        sel = self.tree.selection()
        if not sel:
            return
        midi = int(self.tree.item(sel[0], "values")[1])
        layer.cells.pop(midi, None)
        self._refresh_all()

    def _on_select_layer(self) -> None:
        sel = self.lst_layers.curselection()
        if not sel:
            return
        layer = self.project.layers[sel[0]]
        self.current_layer_id = layer.layer_id
        self._refresh_grid()
        self._refresh_warnings()

    def _refresh_all(self) -> None:
        self._sync_project_meta()
        self.lst_layers.delete(0, "end")
        for layer in self.project.layers:
            self.lst_layers.insert("end", f"{layer.display_name()}  [{len(layer.cells)} notes]")
        if self.current_layer_id:
            for i, layer in enumerate(self.project.layers):
                if layer.layer_id == self.current_layer_id:
                    self.lst_layers.selection_set(i)
                    self.lst_layers.see(i)
                    break
        self._refresh_grid()
        self._refresh_warnings()

    def _refresh_grid(self) -> None:
        self.tree.delete(*self.tree.get_children())
        layer = self._current_layer()
        if not layer:
            return
        warn_midis = {f.midi for f in audit_project(self.project) if f.layer_id == layer.layer_id and f.midi is not None}
        for pitch in chromatic_range(layer.range_low, layer.range_high):
            cell = layer.cells.get(pitch.midi)
            if cell is None:
                self.tree.insert("", "end", values=(pitch.label, pitch.midi, "", "", "", "", "", ""))
                continue
            tags = [cell.origin if cell.origin in ORIGINS else "modelled"]
            if pitch.midi in warn_midis or cell.reporting_status == "above_measured_ceiling":
                tags.append("warn")
            if cell.value <= 0:
                tags = ["bad"]
            pi = ""
            if cell.pi95_low is not None and cell.pi95_high is not None:
                pi = f"{cell.pi95_low:.3g}–{cell.pi95_high:.3g}"
            self.tree.insert(
                "",
                "end",
                values=(
                    cell.note_label,
                    cell.midi,
                    f"{cell.value:.6g}",
                    cell.origin,
                    cell.reporting_status,
                    pi,
                    "; ".join(cell.flags),
                    cell.comment,
                ),
                tags=tuple(tags),
            )

    def _refresh_warnings(self) -> None:
        self.warn.delete(*self.warn.get_children())
        for flag in audit_project(self.project):
            self.warn.insert(
                "",
                "end",
                values=(
                    flag.severity,
                    flag.code,
                    flag.note,
                    flag.approach,
                    flag.message,
                    flag.recommendation,
                ),
                tags=(flag.severity,),
            )

    @staticmethod
    def _flag_text(flags) -> str:
        if not flags:
            return "No extra flags."
        return "\n\n".join(f"[{f.severity}] {f.message}\n→ {f.recommendation}" for f in flags)

    def _open_manual_file(self) -> None:
        path = manual_path()
        if not path.exists():
            messagebox.showerror(
                "Manual not found",
                f"Expected:\n{path}\n\nKeep STE_Lab_User_and_Technical_Manual.html "
                "next to launch_ste_lab.py.",
            )
            return
        _open_local_html(path)

    def _open_help(self) -> None:
        if getattr(self, "_help_win", None) is not None and self._help_win.winfo_exists():
            self._help_win.lift()
            self._help_win.focus_force()
            return
        win = tk.Toplevel(self)
        win.title("STE Lab — Help")
        win.geometry("640x560")
        win.minsize(480, 400)
        self._help_win = win
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Label(bar, text="Short guide  ·  F1 opens this window", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(bar, text="Open full manual", style="Accent.TButton", command=self._open_manual_file).pack(
            side="right", padx=4
        )
        txt = tk.Text(win, wrap="word", font=("Segoe UI", 10), padx=12, pady=10, bg="#f7f4ee")
        scroll = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", _HELP_SHORT)
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)


_HELP_SHORT = """STE Lab in one page

What this program does
You paste measured F-061 Combined Density Metric numbers for one playing
technique. The lab can fill missing notes, estimate another technique
(mute, ponticello, tasto, harmonics), map a curve onto a cousin instrument, or
average two libraries. Invented numbers are tagged modelled — never measured.
F-061 is a project-defined spectral-density number, not “brightness” or
“how many harmonics you hear.”

The four tabs
1. Identity — create a layer (instrument × collection × technique × dynamic).
2. Insert data — type one note, or paste a block into Block paste.
3. Extrapolate — fill, technique transfer (log-ratio), family transfer, Media.
4. Export — write a Zenodo-style Excel, or import one.

Paste rules (the usual mistake)
• Two columns: note and value.  Example:  G3    33.42
• Values only: set “Start note if the paste is only numbers”.
• A bare 72 is a value, not MIDI 72. Write midi72 for the pitch.
• Parse clipboard = read the copy buffer. Parse box and place = commit.
• Clear empties the box, not the layer.

Technique transfer
Anchors are three numbers per line: note, source CDM, target CDM.
The formula is target = source × exp(L), with L learned on the overlap
and held at the last pair beyond that overlap (no dive to zero).

The one rule
Never relabel a filled, transferred, or blended cell as measured.
Empirical_ORCH (measured Orchidea cells) is the principal effect evidence.
Production Media is family-specific: strings average IOWA/ORCH; woodwind
ordinario follows calibration.yaml (default empirical_only). Shared L is
not replication. Sul tasto and inherited pp/ff stay off Empirical_ORCH.
PCHIP fill needs three known notes; two-point PCHIP is rejected.
Violin notes above E7 (MIDI 100) stay flagged and out of summaries.
Harmonics: do not publish modelled values below G5.

Full story
F1 / Help opens this short text. “Open full manual” opens the HTML
in this folder as a local Windows file (not a file:// search in Edge):
STE_Lab_User_and_Technical_Manual.html
It has a layman’s explanation beside the scientific one, recipes, a file map,
and troubleshooting.

Analysis of the five violin Zenodo books
D:\\CORDAS_2\\VIOLIN\\Violino - análise_v2
Editable charts: violin_zenodo_comparacao_editavel.xlsx
"""


def main() -> None:
    app = STEApp()
    app.mainloop()


if __name__ == "__main__":
    main()
