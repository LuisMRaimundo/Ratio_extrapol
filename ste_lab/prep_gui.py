"""Small window: run all STE effects from one preparatory Excel."""

from __future__ import annotations

import os
import queue
import sys
import threading
import webbrowser
from io import StringIO
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .catalog import FAMILIES, instrument_choices, resolve_instrument

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


APP_TITLE = "STE from preparatory Excel"


class _GuiLog(StringIO):
    def __init__(self, put) -> None:
        super().__init__()
        self._put = put

    def write(self, s: str) -> int:
        if s:
            self._put(s)
        return len(s or "")

    def flush(self) -> None:
        return None


class PrepRunnerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        from run_ste_from_preparatory import DEFAULT_OUT, DEFAULT_PREP

        self.title(APP_TITLE)
        self.geometry("780x700")
        self.minsize(680, 580)
        self._busy = False
        self._q: queue.Queue[str] = queue.Queue()
        self._help_win = None
        self._build_style()
        self._build_ui(DEFAULT_PREP, DEFAULT_OUT)
        self.after(120, self._drain_log)

    def _build_style(self) -> None:
        self.configure(bg="#eef2f6")
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#eef2f6")
        style.configure("TLabelframe", background="#eef2f6")
        style.configure("TLabelframe.Label", background="#eef2f6", font=("Segoe UI", 9, "bold"))
        style.configure("TLabel", background="#eef2f6")
        style.configure("TCheckbutton", background="#eef2f6")
        style.configure("Header.TLabel", background="#1f4e79", foreground="white", font=("Segoe UI", 14, "bold"))
        style.configure("Sub.TLabel", background="#1f4e79", foreground="#d6e4f0", font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 9), padding=4)
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

    def _build_ui(self, default_prep: Path, default_out: Path) -> None:
        header = tk.Frame(self, bg="#1f4e79")
        header.pack(fill="x")
        titles = tk.Frame(header, bg="#1f4e79")
        titles.pack(side="left", fill="x", expand=True)
        ttk.Label(titles, text=APP_TITLE, style="Header.TLabel").pack(anchor="w", padx=16, pady=(10, 0))
        ttk.Label(
            titles,
            text="One Excel drives every grounded effect. The main Lab is not required.",
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

        paths = ttk.LabelFrame(self, text="Files")
        paths.pack(fill="x", padx=12, pady=(10, 6))
        paths.columnconfigure(1, weight=1)

        self.var_prep = tk.StringVar(value=str(default_prep))
        self.var_out = tk.StringVar(value=str(default_out))
        ttk.Label(paths, text="Preparatory Excel").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(paths, textvariable=self.var_prep).grid(row=0, column=1, sticky="ew", padx=4, pady=6)
        ttk.Button(paths, text="Browse…", command=self._browse_prep).grid(row=0, column=2, padx=6, pady=6)
        ttk.Label(paths, text="Save folder").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(paths, textvariable=self.var_out).grid(row=1, column=1, sticky="ew", padx=4, pady=6)
        ttk.Button(paths, text="Browse…", command=self._browse_out).grid(row=1, column=2, padx=6, pady=6)

        opts = ttk.LabelFrame(self, text="Session (same labels as the main Lab)")
        opts.pack(fill="x", padx=12, pady=6)
        opts.columnconfigure(1, weight=1)
        opts.columnconfigure(3, weight=1)

        families = list(FAMILIES.values())
        self.cmb_family = ttk.Combobox(opts, values=families, width=24, state="readonly")
        self.cmb_family.set(FAMILIES["bowed_strings"])
        bowed = [s.display_name for s in instrument_choices("bowed_strings")]
        self.cmb_instr = ttk.Combobox(opts, values=bowed, width=24, state="readonly")
        self.cmb_instr.set("Viola")
        self.cmb_family.bind("<<ComboboxSelected>>", lambda e: self._filter_instruments())
        self.cmb_instr.bind("<<ComboboxSelected>>", lambda e: self._on_instrument())

        self.var_operator = tk.StringVar()
        self.var_notes = tk.StringVar()
        ttk.Label(opts, text="Family").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.cmb_family.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Label(opts, text="Instrument").grid(row=0, column=2, sticky="w", padx=6, pady=4)
        self.cmb_instr.grid(row=0, column=3, sticky="ew", padx=4, pady=4)
        ttk.Label(opts, text="Operator").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.var_operator).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Label(opts, text="Session notes").grid(row=1, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(opts, textvariable=self.var_notes).grid(row=1, column=3, sticky="ew", padx=4, pady=4)

        self.var_rebuild = tk.BooleanVar(value=False)
        self.var_open = tk.BooleanVar(value=True)
        from build_ste_preparatory import DEPOSIT_DEFAULTS

        self.var_rebuild_root = tk.StringVar(value=str(DEPOSIT_DEFAULTS["viola"]["root"]))
        flags = ttk.Frame(self)
        flags.pack(fill="x", padx=12, pady=(0, 6))
        flags.columnconfigure(1, weight=1)
        ttk.Checkbutton(
            flags,
            text="Rebuild the preparatory workbook first",
            variable=self.var_rebuild,
            command=self._toggle_rebuild_folder,
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=0, pady=(0, 2))
        self.lbl_rebuild = ttk.Label(flags, text="Look in folder")
        self.ent_rebuild = ttk.Entry(flags, textvariable=self.var_rebuild_root)
        self.btn_rebuild = ttk.Button(flags, text="Browse…", command=self._browse_rebuild)
        self.lbl_rebuild.grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.ent_rebuild.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        self.btn_rebuild.grid(row=1, column=2, padx=6, pady=4)
        ttk.Checkbutton(flags, text="Open the save folder when finished", variable=self.var_open).grid(
            row=2, column=0, columnspan=3, sticky="w"
        )
        self._toggle_rebuild_folder()

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=4)
        self.btn_run = ttk.Button(btns, text="Run all effects", style="Accent.TButton", command=self._run)
        self.btn_run.pack(side="left")
        ttk.Button(btns, text="Open save folder", command=self._open_out).pack(side="left", padx=6)
        ttk.Label(btns, text=f"STE Lab v{__version__}  ·  F1 = Help").pack(side="right")

        log_box = ttk.LabelFrame(self, text="Log")
        log_box.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.txt = tk.Text(log_box, wrap="word", font=("Consolas", 9), height=12, bg="#f7f4ee")
        scroll = ttk.Scrollbar(log_box, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True, padx=6, pady=6)
        self._append(
            "Choose the preparatory Excel and a save folder, then Run all effects.\n"
            "Rebuild looks in the folder under that checkbox — not always viola.\n"
            "Effects and anchors come from the workbook. Sul tasto is not invented.\n"
        )

    def _family_id(self, display: str) -> str:
        for key, label in FAMILIES.items():
            if label == display:
                return key
        return "other"

    def _filter_instruments(self) -> None:
        names = [s.display_name for s in instrument_choices(self._family_id(self.cmb_family.get()))]
        self.cmb_instr["values"] = names
        if names:
            self.cmb_instr.set(names[0])
        self._on_instrument()

    def _on_instrument(self) -> None:
        spec = resolve_instrument(self.cmb_instr.get())
        if not spec:
            return
        from build_ste_preparatory import deposit_defaults, known_out_paths, known_prep_paths, known_roots

        hint = deposit_defaults(spec.instrument_id)
        if not hint:
            return
        raw_root = self.var_rebuild_root.get().strip()
        cur_root = Path(raw_root) if raw_root else None
        if cur_root is None or cur_root in known_roots():
            self.var_rebuild_root.set(str(hint["root"]))
        cur_prep = Path(self.var_prep.get().strip()) if self.var_prep.get().strip() else None
        if cur_prep is None or cur_prep in known_prep_paths():
            self.var_prep.set(str(hint["root"] / hint["prep_name"]))
        cur_out = Path(self.var_out.get().strip()) if self.var_out.get().strip() else None
        if cur_out is None or cur_out in known_out_paths():
            self.var_out.set(str(hint["root"] / hint["out_name"]))

    def _toggle_rebuild_folder(self) -> None:
        state = "normal" if self.var_rebuild.get() else "disabled"
        self.ent_rebuild.configure(state=state)
        self.btn_rebuild.configure(state=state)
        if self.var_rebuild.get() and not self.var_rebuild_root.get().strip():
            self._on_instrument()

    def _browse_rebuild(self) -> None:
        initial = Path(self.var_rebuild_root.get().strip() or ".")
        folder = str(initial if initial.is_dir() else initial.parent)
        path = filedialog.askdirectory(
            title="Look here for Zenodo arco book and research trees",
            initialdir=folder or None,
        )
        if path:
            self.var_rebuild_root.set(path)

    def _browse_prep(self) -> None:
        initial = Path(self.var_prep.get().strip() or ".")
        folder = str(initial.parent if initial.suffix else initial) if initial.exists() else ""
        path = filedialog.askopenfilename(
            title="Open preparatory Excel",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
            initialdir=folder or None,
        )
        if path:
            self.var_prep.set(path)

    def _browse_out(self) -> None:
        initial = Path(self.var_out.get().strip() or ".")
        folder = str(initial if initial.is_dir() else initial.parent)
        path = filedialog.askdirectory(title="Save STE / Zenodo books here", initialdir=folder or None)
        if path:
            self.var_out.set(path)

    def _open_out(self) -> None:
        folder = Path(self.var_out.get().strip())
        if not folder.exists():
            messagebox.showinfo("Folder", f"Not found yet:\n{folder}")
            return
        os.startfile(folder)

    def _append(self, text: str) -> None:
        self.txt.insert("end", text)
        self.txt.see("end")

    def _drain_log(self) -> None:
        while True:
            try:
                chunk = self._q.get_nowait()
            except queue.Empty:
                break
            self._append(chunk)
        self.after(120, self._drain_log)

    def _run(self) -> None:
        if self._busy:
            return
        prep = Path(self.var_prep.get().strip())
        out = Path(self.var_out.get().strip())
        if not self.var_prep.get().strip():
            messagebox.showerror("Missing file", "Choose a preparatory Excel.")
            return
        if not self.var_out.get().strip():
            messagebox.showerror("Missing folder", "Choose a save folder.")
            return
        if not self.var_rebuild.get() and not prep.exists():
            messagebox.showerror("Missing file", f"Preparatory workbook not found:\n{prep}")
            return
        rebuild_root = Path(self.var_rebuild_root.get().strip()) if self.var_rebuild.get() else None
        if self.var_rebuild.get():
            if not self.var_rebuild_root.get().strip():
                messagebox.showerror("Missing folder", "Choose the folder to look in (under Rebuild).")
                return
            if rebuild_root is None or not rebuild_root.exists():
                messagebox.showerror("Missing folder", f"Rebuild look-in folder not found:\n{rebuild_root}")
                return
        spec = resolve_instrument(self.cmb_instr.get())
        instrument = spec.instrument_id if spec else self.cmb_instr.get()
        self._busy = True
        self.btn_run.state(["disabled"])
        self._append(f"\n--- starting {prep.name} → {out} ---\n")

        def work() -> None:
            nonlocal prep
            old = sys.stdout
            sys.stdout = _GuiLog(self._q.put)
            err = ""
            written = []
            rebuild_warnings: list[str] = []
            try:
                from ste_lab.pipeline import pipeline_for_instrument, run_selected_pipeline

                print(f"Pipeline  {pipeline_for_instrument(instrument)}  instrument={instrument}")
                if not self.var_rebuild.get() and not prep.exists():
                    raise FileNotFoundError(f"missing preparatory workbook: {prep}")
                written = run_selected_pipeline(
                    instrument,
                    prep.resolve() if prep.exists() or not self.var_rebuild.get() else prep,
                    out.resolve(),
                    rebuild=bool(self.var_rebuild.get() and rebuild_root is not None),
                    rebuild_root=rebuild_root,
                    operator=self.var_operator.get(),
                    notes=self.var_notes.get(),
                )
                import build_ste_preparatory as bsp

                if bsp.last_path:
                    prep = bsp.last_path
                    self.after(0, lambda p=str(prep): self.var_prep.set(p))
                rebuild_warnings = list(bsp.last_warnings or [])
                print("\nDone:")
                for path in written:
                    print(" ", path)
            except (Exception, SystemExit) as exc:
                err = str(exc) or exc.__class__.__name__
                print(f"\nStopped: {err}")
            finally:
                sys.stdout = old
                self.after(
                    0,
                    lambda e=err, w=written, o=out, rw=list(rebuild_warnings): self._finished(
                        e, w, o, rw
                    ),
                )

        threading.Thread(target=work, daemon=True).start()

    def _finished(self, err: str, written: list, out: Path, rebuild_warnings: list | None = None) -> None:
        self._busy = False
        self.btn_run.state(["!disabled"])
        if err:
            messagebox.showerror("STE from preparatory", err)
            return
        if rebuild_warnings:
            preview = "\n".join(rebuild_warnings[:12])
            extra = "" if len(rebuild_warnings) <= 12 else f"\n… and {len(rebuild_warnings) - 12} more (see log)."
            messagebox.showwarning(
                "Research folders",
                "Rebuild continued, but some expected folders or compiled books were missing:\n\n"
                + preview
                + extra,
            )
        messagebox.showinfo(
            "STE from preparatory",
            f"Wrote {len(written)} workbook(s) to:\n{out}",
        )
        if self.var_open.get() and out.exists():
            os.startfile(out)

    def _open_manual_file(self) -> None:
        path = manual_path()
        if not path.exists():
            messagebox.showerror(
                "Manual not found",
                f"Expected:\n{path}\n\nKeep {MANUAL_NAME} next to launch_ste_lab.py.",
            )
            return
        _open_local_html(path)

    def _open_help(self) -> None:
        if self._help_win is not None and self._help_win.winfo_exists():
            self._help_win.lift()
            self._help_win.focus_force()
            return
        win = tk.Toplevel(self)
        win.title("STE from preparatory — Help")
        win.geometry("620x420")
        win.minsize(460, 320)
        self._help_win = win
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Label(bar, text="This window only  ·  F1", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(bar, text="Open full Lab manual", style="Accent.TButton", command=self._open_manual_file).pack(
            side="right", padx=4
        )
        txt = tk.Text(win, wrap="word", font=("Segoe UI", 10), padx=12, pady=10, bg="#f7f4ee")
        scroll = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", _HELP)
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)


_HELP = """STE from one preparatory Excel

What to point at
Preparatory Excel — the workbook with Paste_arco, Paste_effects_mf and
Anchors_all (for example Viola_STE_preparatory.xlsx).
Save folder — where the STE and Zenodo books are written.

What the other boxes do
Family / Instrument, Operator and Session notes are the same labels as
on the main STE Lab. They are printed on the exported workbooks. The
effects themselves still come from the Excel, not from these boxes.
Without Rebuild, Instrument = Cello with a viola Excel still writes
viola numbers under Cello_STE_* names. Do not deposit those.

Rebuild first
Tick Rebuild, then set Look in folder to the deposit you want
(e.g. D:\\CORDAS_3\\CELLO, D:\\CORDAS_3\\VIOLA 4, or
D:\\MADEIRAS_2\\BASS_CLARINET). Changing Instrument suggests that
folder and selects the family pipeline (strings vs woodwinds).
Rebuild writes {Instrument}_STE_preparatory.xlsx there, then runs.
Strings: Iowa comes from the Zenodo book (by-string sheets), not from
_Sustains_Stable. Woodwinds: Iowa comes from target compiled trees
(merged; longest-file-wins is prohibited) and generated media is
written under generated/, never over the source workbook.
Orchidea / Phil / McGill still use compiled research books under
_Sustains_Stable. Missing folders produce a WARNING; the run continues.
If a string deposit has effect trees and Anchors_all is empty, the run
stops (DataDiscoveryError) instead of writing only ordinario.

Orchidea recordings (strings)
If Orchidea recorded that effect at that dynamic, those cells stay
measured. L is not applied to that layer. L is only for IOWA, or for
a dynamic Orchidea did not record. Woodwind ordinario uses two_log_ratio
family transfer + empirical_only; it does not invent string techniques.

Evidence_Map
principal_evidence is Orchidea measured only. Iowa that used the same
dynamic’s L is modelled, not inherited. Philharmonia and McGill are
context, never principal.

What it will not do
It does not invent sul tasto. It does not open the big Lab. Technique,
collection and anchors stay in the Excel.

Command line (no window)
python run_ste_from_preparatory.py --rebuild-prep --rebuild-root "D:\\CORDAS_3\\CELLO" --instrument cello
python -m pytest test_pipeline_regression.py test_calibration.py test_run_ste_from_preparatory.py

Full story
F1 / Help opens this short text. “Open full Lab manual” opens the
HTML in this folder as a local file (not an Edge search):
STE_Lab_User_and_Technical_Manual.html
"""


def main() -> None:
    app = PrepRunnerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
