import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import DRIVERS, FAMILIES, FLOWS, build_plan
from .exporter import ExportError, export_etl_to_csv
from .runner import CollectorError, run_config, start_collection, stop_collection


class DriverCollectorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Driver Collector")
        self.root.geometry("980x680")
        self.root.minsize(840, 560)
        self.family_var = tk.StringVar(value=FAMILIES[0])
        self.driver_var = tk.StringVar(value=DRIVERS[0])
        self.flow_var = tk.StringVar(value=FLOWS[0])
        self.tag_var = tk.StringVar(value="")
        self.symbols_var = tk.StringVar(value="")
        self.etl_var = tk.StringVar(value="")
        self.csv_var = tk.StringVar(value="")
        self.create_widgets()
        self.refresh_plan()

    def create_widgets(self):
        root = ttk.Frame(self.root, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(root, text="Driver Collector", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        ttk.Label(root, text="Driver family").grid(row=1, column=0, sticky="w")
        ttk.Combobox(root, textvariable=self.family_var, values=FAMILIES, state="readonly").grid(row=2, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(root, text="Driver type").grid(row=1, column=1, sticky="w")
        ttk.Combobox(root, textvariable=self.driver_var, values=DRIVERS, state="readonly").grid(row=2, column=1, sticky="ew", padx=(0, 8))
        ttk.Label(root, text="Flow").grid(row=1, column=2, sticky="w")
        ttk.Combobox(root, textvariable=self.flow_var, values=FLOWS, state="readonly").grid(row=2, column=2, sticky="ew", padx=(0, 8))
        ttk.Button(root, text="Show Plan", command=self.refresh_plan).grid(row=2, column=3, sticky="ew")

        for variable in (self.family_var, self.driver_var, self.flow_var):
            variable.trace_add("write", lambda *_args: self.refresh_plan())

        ttk.Label(root, text="Optional tag/session suffix").grid(row=3, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(root, textvariable=self.tag_var).grid(row=4, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(root, text="Symbols / PDB path").grid(row=3, column=1, sticky="w", pady=(12, 0))
        ttk.Entry(root, textvariable=self.symbols_var).grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(root, text="Load Symbols", command=self.load_symbols).grid(row=4, column=3, sticky="ew")

        buttons = ttk.Frame(root)
        buttons.grid(row=5, column=0, columnspan=4, sticky="ew", pady=12)
        ttk.Button(buttons, text="Run Verbose Config", command=self.run_verbose_config).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Start Collection", command=self.start).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Stop Collection", command=self.stop).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Separator(root).grid(row=6, column=0, columnspan=4, sticky="ew", pady=(4, 12))

        ttk.Label(root, text="ETL file").grid(row=7, column=0, sticky="w")
        ttk.Entry(root, textvariable=self.etl_var).grid(row=8, column=0, columnspan=3, sticky="ew", padx=(0, 8))
        ttk.Button(root, text="Browse ETL", command=self.browse_etl).grid(row=8, column=3, sticky="ew")

        ttk.Label(root, text="CSV output").grid(row=9, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(root, textvariable=self.csv_var).grid(row=10, column=0, columnspan=3, sticky="ew", padx=(0, 8))
        ttk.Button(root, text="Export CSV", command=self.export_csv).grid(row=10, column=3, sticky="ew")

        self.output = tk.Text(root, wrap=tk.WORD, font=("Consolas", 10), height=18)
        yscroll = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.output.yview)
        self.output.configure(yscrollcommand=yscroll.set)
        self.output.grid(row=11, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        yscroll.grid(row=11, column=3, sticky="nsw", pady=(12, 0))

        for column in range(4):
            root.grid_columnconfigure(column, weight=1)
        root.grid_rowconfigure(11, weight=1)

    def current_plan(self):
        return build_plan(self.family_var.get(), self.driver_var.get(), self.flow_var.get())

    def refresh_plan(self):
        self.write_output(self.current_plan().describe())

    def load_symbols(self):
        path = filedialog.askopenfilename(title="Select PDB or symbol file", filetypes=[("PDB files", "*.pdb"), ("All files", "*.*")])
        if not path:
            path = filedialog.askdirectory(title="Select symbol folder")
        if path:
            self.symbols_var.set(path)
            self.write_output(f"Symbols loaded for export child tools:\n{path}")

    def run_verbose_config(self):
        self.run_action("Verbose Config", lambda: run_config(self.current_plan()))

    def start(self):
        self.run_action("Start Collection", lambda: start_collection(self.current_plan(), tag=self.tag_var.get().strip()))

    def stop(self):
        self.run_action("Stop Collection", lambda: stop_collection(self.current_plan()))

    def browse_etl(self):
        path = filedialog.askopenfilename(title="Select ETL file", filetypes=[("ETL files", "*.etl"), ("All files", "*.*")])
        if not path:
            return
        self.etl_var.set(path)
        if not self.csv_var.get().strip():
            self.csv_var.set(str(Path(path).with_suffix(".csv")))

    def export_csv(self):
        etl = self.etl_var.get().strip()
        output = self.csv_var.get().strip()
        if not etl or not output:
            messagebox.showinfo("Export CSV", "Select an ETL file and CSV output path first.")
            return
        self.run_action("Export CSV", lambda: f"Wrote CSV: {export_etl_to_csv(etl, output, self.symbols_var.get().strip() or None)}")

    def run_action(self, title, action):
        try:
            result = action()
        except (CollectorError, ExportError, OSError, ValueError) as exc:
            self.write_output(f"{title} failed:\n{exc}")
            return
        self.write_output(f"{title} complete:\n{result}")

    def write_output(self, text):
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, text)

    def run(self):
        self.root.mainloop()


def launch_gui():
    DriverCollectorApp().run()
