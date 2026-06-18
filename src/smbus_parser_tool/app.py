import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .parser import parse_csv, summarize


class SmbusParserApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SMBus Parser Tool")
        self.root.geometry("860x560")
        self.root.minsize(720, 460)
        self.path_var = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        header = ttk.Frame(self.root, padding=12)
        header.pack(fill=tk.X)
        ttk.Label(header, text="SMBus Parser Tool", font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT)
        ttk.Button(header, text="Open CSV", command=self.open_csv).pack(side=tk.RIGHT)

        path_frame = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        path_frame.pack(fill=tk.X)
        ttk.Entry(path_frame, textvariable=self.path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="Analyze", command=self.analyze_current).pack(side=tk.RIGHT, padx=(8, 0))

        self.output = tk.Text(self.root, wrap=tk.WORD, font=("Consolas", 10))
        scroll = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.output.yview)
        self.output.configure(yscrollcommand=scroll.set)
        self.output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0), pady=(0, 12))
        scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 12), pady=(0, 12))

    def open_csv(self):
        path = filedialog.askopenfilename(title="Open SMBus CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.path_var.set(path)
            self.analyze_current()

    def analyze_current(self):
        path = self.path_var.get().strip()
        if not path:
            messagebox.showinfo("SMBus Parser Tool", "Select a CSV file first.")
            return
        try:
            records = parse_csv(path)
        except Exception as exc:
            messagebox.showerror("Parse Error", str(exc))
            return
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, summarize(records))

    def run(self):
        self.root.mainloop()


def launch_gui():
    SmbusParserApp().run()
