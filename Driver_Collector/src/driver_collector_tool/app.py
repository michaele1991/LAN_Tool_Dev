import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import DRIVERS, FAMILIES, FLOWS, build_plan
from .exporter import ExportError, export_etl_to_csv
from .runner import CollectorError, run_config, start_collection, stop_collection

# ── palette ──────────────────────────────────────────────────────────────────
C = {
    "bg":           "#1e1e1e",
    "surface":      "#252526",
    "surface2":     "#2d2d30",
    "border":       "#3c3c3c",
    "text":         "#d4d4d4",
    "text_dim":     "#888888",
    "text_bright":  "#ffffff",
    "accent":       "#007acc",
    "accent_hover": "#1c97ea",
    "success":      "#4ec9b0",
    "warning":      "#dcdcaa",
    "danger":       "#f44747",
    "green_btn":    "#1e7b34",
    "red_btn":      "#9b1c1c",
    "blue_btn":     "#0e639c",
    "amber_btn":    "#7a5c00",
}


def _hex(color: str) -> str:
    return color


class DriverCollectorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Intel LAN Driver Collector")
        self.root.geometry("1080x760")
        self.root.minsize(900, 620)
        self.root.configure(bg=C["bg"])

        self.family_var = tk.StringVar(value=FAMILIES[0])
        self.driver_var = tk.StringVar(value=DRIVERS[0])
        self.flow_var   = tk.StringVar(value=FLOWS[0])
        self.tag_var     = tk.StringVar(value="")
        self.symbols_var = tk.StringVar(value="")
        self.etl_var     = tk.StringVar(value="")
        self.csv_var     = tk.StringVar(value="")

        self._configure_ttk_style()
        self._create_header()
        self._create_body()
        self._create_statusbar()

        for v in (self.family_var, self.driver_var, self.flow_var):
            v.trace_add("write", lambda *_: self.refresh_plan())

        self.refresh_plan()

    # ── style ─────────────────────────────────────────────────────────────────
    def _configure_ttk_style(self):
        st = ttk.Style(self.root)
        st.theme_use("clam")
        st.configure(".",
            background=C["bg"], foreground=C["text"],
            fieldbackground=C["surface2"], selectbackground=C["accent"],
            selectforeground=C["text_bright"], troughcolor=C["surface2"],
            bordercolor=C["border"], darkcolor=C["surface"], lightcolor=C["surface"],
            font=("Segoe UI", 10))
        st.configure("TFrame", background=C["bg"])
        st.configure("Card.TFrame", background=C["surface"], relief="flat")
        st.configure("TLabel", background=C["bg"], foreground=C["text"], font=("Segoe UI", 10))
        st.configure("Header.TLabel", background=C["surface"], foreground=C["text_bright"], font=("Segoe UI", 10))
        st.configure("SectionTitle.TLabel", background=C["surface"], foreground=C["accent"],
                     font=("Segoe UI", 9, "bold"))
        st.configure("TCombobox", fieldbackground=C["surface2"], background=C["surface2"],
                     foreground=C["text"], selectbackground=C["surface2"],
                     arrowcolor=C["text_dim"], insertcolor=C["text"])
        st.map("TCombobox", fieldbackground=[("readonly", C["surface2"])],
               selectbackground=[("readonly", C["surface2"])],
               selectforeground=[("readonly", C["text"])])
        st.configure("TEntry", fieldbackground=C["surface2"], foreground=C["text"],
                     insertcolor=C["text"])
        st.configure("TSeparator", background=C["border"])
        st.configure("Vertical.TScrollbar", background=C["surface2"],
                     troughcolor=C["bg"], arrowcolor=C["text_dim"], bordercolor=C["border"])

    # ── header ────────────────────────────────────────────────────────────────
    def _create_header(self):
        hdr = tk.Frame(self.root, bg=C["surface"], height=56)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        # coloured left accent bar
        tk.Frame(hdr, bg=C["accent"], width=4).pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(hdr, text="Intel LAN  ·  Driver Collector",
                 font=("Segoe UI", 16, "bold"),
                 bg=C["surface"], fg=C["text_bright"]).pack(side=tk.LEFT, padx=14, pady=10)

        self._status_dot = tk.Label(hdr, text="●", font=("Segoe UI", 12),
                                     bg=C["surface"], fg=C["text_dim"])
        self._status_dot.pack(side=tk.RIGHT, padx=(0, 8))
        self._status_lbl = tk.Label(hdr, text="Idle",
                                     font=("Segoe UI", 9), bg=C["surface"], fg=C["text_dim"])
        self._status_lbl.pack(side=tk.RIGHT)

    # ── body ──────────────────────────────────────────────────────────────────
    def _create_body(self):
        paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL,
                               bg=C["border"], sashwidth=3, sashrelief=tk.FLAT)
        paned.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        left  = tk.Frame(paned, bg=C["bg"])
        right = tk.Frame(paned, bg=C["bg"])
        paned.add(left,  minsize=440)
        paned.add(right, minsize=380)

        self._create_selection_card(left)
        self._create_options_card(left)
        self._create_actions_card(left)
        self._create_export_card(left)
        self._create_output_pane(right)

    # ── section card helper ───────────────────────────────────────────────────
    def _card(self, parent, title: str, top_pad: int = 10) -> tk.Frame:
        outer = tk.Frame(parent, bg=C["bg"])
        outer.pack(fill=tk.X, padx=10, pady=(top_pad, 0))

        header_row = tk.Frame(outer, bg=C["surface"], height=28)
        header_row.pack(fill=tk.X)
        header_row.pack_propagate(False)
        tk.Frame(header_row, bg=C["accent"], width=3).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(header_row, text=title.upper(), font=("Segoe UI", 8, "bold"),
                 bg=C["surface"], fg=C["accent"]).pack(side=tk.LEFT, padx=8, pady=4)

        body = tk.Frame(outer, bg=C["surface"], padx=10, pady=10)
        body.pack(fill=tk.X)
        return body

    def _lbl(self, parent, text: str) -> tk.Label:
        return tk.Label(parent, text=text, font=("Segoe UI", 9),
                        bg=C["surface"], fg=C["text_dim"])

    def _combo(self, parent, variable, values) -> ttk.Combobox:
        cb = ttk.Combobox(parent, textvariable=variable, values=values,
                          state="readonly", font=("Segoe UI", 10))
        return cb

    def _entry(self, parent, variable, placeholder: str = "") -> ttk.Entry:
        if placeholder:
            # Do NOT bind textvariable — keep StringVar clean (always "" when
            # placeholder is showing).  Storing the placeholder in the StringVar
            # causes list2cmdline to quote it (it has spaces), so bat files receive
            # `"e.g. test_001"` as %4, which breaks `if "%session_name%"` in cmd.
            e = ttk.Entry(parent, font=("Consolas", 9))
            _active = [False]   # True = user content shown, False = placeholder shown

            def _show_placeholder():
                e.delete(0, tk.END)
                e.insert(0, placeholder)
                e.configure(foreground=C["text_dim"])
                _active[0] = False

            def _on_focus_in(_evt):
                if not _active[0]:
                    e.delete(0, tk.END)
                    e.configure(foreground=C["text"])
                    _active[0] = True

            def _on_focus_out(_evt):
                val = e.get().strip()
                variable.set(val)
                if not val:
                    _show_placeholder()

            e.bind("<FocusIn>",  _on_focus_in)
            e.bind("<FocusOut>", _on_focus_out)
            variable.set("")
            _show_placeholder()
        else:
            e = ttk.Entry(parent, textvariable=variable, font=("Consolas", 9))
        return e

    def _btn(self, parent, text: str, command, bg: str = C["blue_btn"],
             fg: str = C["text_bright"], icon: str = "") -> tk.Button:
        label = f"{icon}  {text}" if icon else text
        b = tk.Button(parent, text=label, command=command,
                      bg=bg, fg=fg, activebackground=C["accent_hover"],
                      activeforeground=C["text_bright"], relief=tk.FLAT,
                      font=("Segoe UI", 9, "bold"), padx=12, pady=6,
                      cursor="hand2", bd=0)
        b.bind("<Enter>", lambda _e: b.configure(bg=C["accent_hover"]))
        b.bind("<Leave>", lambda _e: b.configure(bg=bg))
        return b

    # ── section: target selection ─────────────────────────────────────────────
    def _create_selection_card(self, parent):
        body = self._card(parent, "1  ·  Target Selection")

        cols = tk.Frame(body, bg=C["surface"])
        cols.pack(fill=tk.X)
        for col in range(3):
            cols.grid_columnconfigure(col, weight=1)

        labels   = ["Driver Family", "Driver Type", "Flow"]
        combos   = [
            (self.family_var, list(FAMILIES)),
            (self.driver_var, list(DRIVERS)),
            (self.flow_var,   list(FLOWS)),
        ]
        self._flow_indicator = []
        for col, (lbl_text, (var, vals)) in enumerate(zip(labels, combos)):
            self._lbl(cols, lbl_text).grid(row=0, column=col, sticky="w", padx=(0, 8))
            cb = self._combo(cols, var, vals)
            cb.grid(row=1, column=col, sticky="ew", padx=(0, 8), pady=(2, 0))
            self._flow_indicator.append(cb)

    # ── section: options ──────────────────────────────────────────────────────
    def _create_options_card(self, parent):
        body = self._card(parent, "2  ·  Options")

        cols = tk.Frame(body, bg=C["surface"])
        cols.pack(fill=tk.X)
        cols.grid_columnconfigure(0, weight=1)
        cols.grid_columnconfigure(1, weight=3)
        cols.grid_columnconfigure(2, weight=0)

        self._lbl(cols, "Session tag  (optional)").grid(row=0, column=0, sticky="w")
        self._entry(cols, self.tag_var, "e.g. test_001").grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self._lbl(cols, "Symbols / PDB path").grid(row=0, column=1, sticky="w")
        self._entry(cols, self.symbols_var, "e.g. C:\\Symbols\\e1dn.pdb").grid(
            row=1, column=1, sticky="ew", padx=(0, 8))

        self._btn(cols, "Browse PDB", self.load_symbols, bg=C["surface2"],
                  fg=C["text"], icon="📂").grid(row=1, column=2, sticky="ew")

    # ── section: actions ──────────────────────────────────────────────────────
    def _create_actions_card(self, parent):
        body = self._card(parent, "3  ·  Collection Control")

        row1 = tk.Frame(body, bg=C["surface"])
        row1.pack(fill=tk.X, pady=(0, 6))

        self._btn(row1, "Verbose Config",   self.run_verbose_config,
                  bg=C["amber_btn"], icon="⚙").pack(side=tk.LEFT, padx=(0, 8))
        self._btn(row1, "Start Collection", self.start,
                  bg=C["green_btn"], icon="▶").pack(side=tk.LEFT, padx=(0, 8))
        self._btn(row1, "Stop Collection",  self.stop,
                  bg=C["red_btn"],   icon="■").pack(side=tk.LEFT)

        # plan preview row
        self._plan_var = tk.StringVar(value="")
        plan_frame = tk.Frame(body, bg=C["surface2"], padx=8, pady=6)
        plan_frame.pack(fill=tk.X, pady=(6, 0))
        tk.Label(plan_frame, textvariable=self._plan_var,
                 font=("Consolas", 8), bg=C["surface2"], fg=C["text_dim"],
                 justify=tk.LEFT, anchor="w").pack(fill=tk.X)

    # ── section: export ───────────────────────────────────────────────────────
    def _create_export_card(self, parent):
        body = self._card(parent, "4  ·  ETL  →  CSV Export")

        grid = tk.Frame(body, bg=C["surface"])
        grid.pack(fill=tk.X)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=0)

        self._lbl(grid, "ETL file").grid(row=0, column=0, sticky="w")
        etl_row = tk.Frame(grid, bg=C["surface"])
        etl_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        etl_row.grid_columnconfigure(0, weight=1)
        self._entry(etl_row, self.etl_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._btn(etl_row, "Browse ETL", self.browse_etl, bg=C["surface2"],
                  fg=C["text"], icon="📂").grid(row=0, column=1)

        self._lbl(grid, "CSV output").grid(row=2, column=0, sticky="w")
        csv_row = tk.Frame(grid, bg=C["surface"])
        csv_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        csv_row.grid_columnconfigure(0, weight=1)
        self._entry(csv_row, self.csv_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._btn(csv_row, "Export CSV", self.export_csv,
                  bg=C["blue_btn"], icon="💾").grid(row=0, column=1)

    # ── output pane ───────────────────────────────────────────────────────────
    def _create_output_pane(self, parent):
        hdr = tk.Frame(parent, bg=C["surface"], height=28)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C["success"], width=3).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(hdr, text="OUTPUT / PLAN", font=("Segoe UI", 8, "bold"),
                 bg=C["surface"], fg=C["success"]).pack(side=tk.LEFT, padx=8, pady=4)

        frame = tk.Frame(parent, bg=C["bg"])
        frame.pack(fill=tk.BOTH, expand=True, padx=(0, 0))
        self.output = tk.Text(
            frame, wrap=tk.WORD,
            font=("Consolas", 10),
            bg=C["surface"], fg=C["text"],
            insertbackground=C["text"],
            selectbackground=C["accent"],
            relief=tk.FLAT, padx=12, pady=10,
            spacing1=2, spacing3=2,
        )
        self.output.tag_configure("key",   foreground=C["warning"])
        self.output.tag_configure("val",   foreground=C["text"])
        self.output.tag_configure("ok",    foreground=C["success"])
        self.output.tag_configure("err",   foreground=C["danger"])
        self.output.tag_configure("note",  foreground=C["text_dim"])
        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.output.yview)
        self.output.configure(yscrollcommand=yscroll.set)
        self.output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ── status bar ────────────────────────────────────────────────────────────
    def _create_statusbar(self):
        bar = tk.Frame(self.root, bg=C["surface"], height=24)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)
        self._statusbar_lbl = tk.Label(bar, text="Ready",
                                       font=("Segoe UI", 8), bg=C["surface"], fg=C["text_dim"])
        self._statusbar_lbl.pack(side=tk.LEFT, padx=10)

    def _set_status(self, text: str, color: str = C["text_dim"], dot: str | None = None):
        self._statusbar_lbl.configure(text=text)
        if dot:
            self._status_dot.configure(fg=dot)
            self._status_lbl.configure(text=text[:40])

    # ── logic ─────────────────────────────────────────────────────────────────
    def current_plan(self):
        return build_plan(self.family_var.get(), self.driver_var.get(), self.flow_var.get())

    def refresh_plan(self):
        plan = self.current_plan()
        # update compact plan preview
        short = f"Start: {plan.start_script.name if plan.start_script else '—'} | Stop: {plan.stop_script.name if plan.stop_script else '—'}"
        if not plan.supported:
            short = f"⚠ {plan.notes[:80]}"
        self._plan_var.set(short)
        self._write_plan(plan)

    def _write_plan(self, plan):
        self.output.configure(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        for line in plan.describe().splitlines():
            if ": " in line:
                key, _, val = line.partition(": ")
                self.output.insert(tk.END, key + ": ", "key")
                tag = "ok" if val.lower() == "yes" else ("err" if val.lower() == "no" else "val")
                self.output.insert(tk.END, val + "\n", tag)
            else:
                self.output.insert(tk.END, line + "\n", "note")
        self.output.configure(state=tk.NORMAL)

    def load_symbols(self):
        path = filedialog.askopenfilename(
            title="Select PDB or symbol file",
            filetypes=[("PDB files", "*.pdb"), ("All files", "*.*")])
        if not path:
            path = filedialog.askdirectory(title="Select symbol folder")
        if path:
            self.symbols_var.set(path)
            self._set_status(f"Symbols: {Path(path).name}", C["success"], dot=C["success"])
            self.write_output("Symbols loaded for export:\n" + path, "ok")

    def run_verbose_config(self):
        self._run_action("Verbose Config", lambda: run_config(self.current_plan()))

    def start(self):
        self._run_action("Start Collection", lambda: start_collection(
            self.current_plan(), tag=self.tag_var.get().strip()))

    def stop(self):
        self._run_action("Stop Collection", lambda: stop_collection(
            self.current_plan(), tag=self.tag_var.get().strip()))

    def browse_etl(self):
        path = filedialog.askopenfilename(
            title="Select ETL file",
            filetypes=[("ETL files", "*.etl"), ("All files", "*.*")])
        if not path:
            return
        self.etl_var.set(path)
        if not self.csv_var.get().strip():
            self.csv_var.set(str(Path(path).with_suffix(".csv")))
        self._set_status(f"ETL: {Path(path).name}")

    def export_csv(self):
        etl = self.etl_var.get().strip()
        output = self.csv_var.get().strip()
        if not etl or not output:
            messagebox.showinfo("Export CSV", "Select an ETL file and CSV output path first.")
            return
        self._run_action("Export CSV", lambda: "Wrote CSV: " + str(
            export_etl_to_csv(etl, output, self.symbols_var.get().strip() or None)))

    def _run_action(self, title: str, action):
        self._set_status(f"{title}…", C["warning"], dot=C["warning"])
        try:
            result = action()
            self.write_output(f"✔  {title}\n\n{result}", "ok")
            self._set_status(f"Done: {title}", C["success"], dot=C["success"])
        except (CollectorError, ExportError, OSError, ValueError) as exc:
            self.write_output(f"✖  {title} failed\n\n{exc}", "err")
            self._set_status(f"Error: {title}", C["danger"], dot=C["danger"])

    def write_output(self, text: str, default_tag: str = "val"):
        self.output.configure(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, text + "\n", default_tag)
        self.output.configure(state=tk.NORMAL)

    def run(self):
        self.root.mainloop()


def launch_gui():
    DriverCollectorApp().run()
