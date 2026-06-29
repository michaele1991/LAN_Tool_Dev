"""
NVM AI Assistant - Intel GBE NVM Configuration Advisor
Powered by Anthropic Claude API

Form-based: pick project / offset / bit / mode, then talk to Claude.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
import threading

try:
    import anthropic
except ImportError:
    anthropic = None

APP_TITLE   = "Intel GBE NVM AI Assistant"
APP_VERSION = "v1.0"
MODEL       = "claude-opus-4-5"
MAX_TOKENS  = 4096
CONFIG_FILE = Path(__file__).parent / "config.json"
NVM_ROOT    = "Eng_GBE_Image"

SYSTEM_PROMPT = """You are an expert Intel GBE (Gigabit Ethernet) NVM (Non-Volatile Memory) configuration engineer.
You help engineers configure NVM settings for Intel Ethernet controllers
(Nahum7, Nahum8, Nahum9, Nahum10, Nahum11, Nahum13, MTL and similar platforms).

Your expertise includes:
- GBE NVM word/bit field definitions and their effects on hardware behavior
- LAN SW (V) vs Non-LAN SW (LM) configuration differences
- Device ID, SKU, silicon stepping considerations
- Wake-on-LAN, PCIe, power management, EEE settings
- Checksum calculation: word 0x3F must equal 0xBABA
- C-Spec and RTL register naming conventions

When answering a configuration question:
1. Confirm the exact word offset (hex) and bit field(s) to modify
2. Specify the exact hex value for V and/or LM with clear rationale
3. Note any dependencies, ordering requirements, or platform-specific caveats
4. Warn about any risk of incorrect configuration

Format register changes as a table when possible:
| Offset | Bits | Field Name | V Value | LM Value | Reason |
"""


def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"api_key": "", "model": MODEL}


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def find_nvm_root():
    """Search upward from this script for the Eng_GBE_Image folder."""
    here = Path(__file__).resolve().parent
    for candidate in [here, here.parent, here.parent.parent, here.parent.parent.parent]:
        p = candidate / NVM_ROOT
        if p.is_dir():
            return p
    return None


class NvmAiApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  {APP_VERSION}")
        self.geometry("1150x820")
        self.minsize(900, 640)
        self.config_data = load_config()
        self.client = None
        self.conversation = []
        self.nvm_root = find_nvm_root()
        self.excel_context = ""
        self._build_ui()
        self._load_projects()
        self._try_init_client()

    def _apply_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("vista")
        except Exception:
            pass
        s.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

    def _build_ui(self):
        self._apply_styles()

        # Header
        hdr = tk.Frame(self, bg="#0071c5", height=50)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  Intel GBE NVM AI Assistant  {APP_VERSION}",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#0071c5").pack(side=tk.LEFT, padx=14, pady=10)
        tk.Label(hdr, text="Powered by Claude (Anthropic)",
                 font=("Segoe UI", 9), fg="#a8d4f0", bg="#0071c5").pack(side=tk.RIGHT, padx=16)

        # Main layout
        main = tk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # LEFT form (fixed 300px)
        form_frame = tk.Frame(main, width=310, bd=1, relief=tk.FLAT)
        form_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        form_frame.pack_propagate(False)
        self._build_form(form_frame)

        # RIGHT chat
        chat_frame = tk.Frame(main)
        chat_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_chat(chat_frame)

        # Status bar
        sb = tk.Frame(self, bg="#e8e8e8", height=22, relief=tk.SUNKEN, bd=1)
        sb.pack(side=tk.BOTTOM, fill=tk.X)
        sb.pack_propagate(False)
        self.api_status_lbl = tk.Label(sb, text="API: not connected", fg="#888",
                                       font=("Segoe UI", 8), bg="#e8e8e8")
        self.api_status_lbl.pack(side=tk.LEFT, padx=8)
        root_txt = f"NVM root: {self.nvm_root}" if self.nvm_root else f"'{NVM_ROOT}' folder not found"
        root_col = "#555" if self.nvm_root else "#cc0000"
        self.nvm_status_lbl = tk.Label(sb, text=root_txt, fg=root_col,
                                       font=("Segoe UI", 8), bg="#e8e8e8")
        self.nvm_status_lbl.pack(side=tk.LEFT, padx=16)

    def _card(self, parent, title):
        f = ttk.LabelFrame(parent, text=title, padding=(8, 4))
        f.pack(fill=tk.X, pady=(0, 8))
        return f

    def _build_form(self, parent):
        tk.Label(parent, text="NVM Configuration",
                 font=("Segoe UI", 11, "bold"), fg="#0071c5").pack(anchor=tk.W, pady=(6, 10))

        # 1. Project
        c = self._card(parent, "1. Project")
        c.columnconfigure(1, weight=1)
        tk.Label(c, text="Platform:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
        self.project_var = tk.StringVar()
        self.project_cb = ttk.Combobox(c, textvariable=self.project_var,
                                       state="readonly", width=24)
        self.project_cb.grid(row=0, column=1, sticky="ew", padx=(6, 0), pady=3)
        self.project_cb.bind("<<ComboboxSelected>>", lambda _: self._on_project_change())
        tk.Label(c, text="Folder:", font=("Segoe UI", 8), fg="#888").grid(row=1, column=0, sticky="w")
        self.folder_lbl = tk.Label(c, text="—", font=("Segoe UI", 8), fg="#0071c5",
                                   wraplength=220, justify=tk.LEFT)
        self.folder_lbl.grid(row=1, column=1, sticky="w", padx=(6, 0))

        # 2. Register
        c2 = self._card(parent, "2. Register")
        c2.columnconfigure(1, weight=1)
        tk.Label(c2, text="Offset (hex):", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
        self.offset_var = tk.StringVar()
        ttk.Entry(c2, textvariable=self.offset_var, width=12).grid(row=0, column=1, sticky="w", padx=(6, 0), pady=3)

        tk.Label(c2, text="Bit(s):", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=3)
        self.bits_var = tk.StringVar()
        ttk.Entry(c2, textvariable=self.bits_var, width=12).grid(row=1, column=1, sticky="w", padx=(6, 0), pady=3)
        tk.Label(c2, text="e.g.  3   or   15:8", font=("Segoe UI", 8), fg="#999").grid(
            row=2, column=0, columnspan=2, sticky="w")

        tk.Label(c2, text="Register name:", font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", pady=(10, 3))
        self.reg_name_var = tk.StringVar()
        ttk.Entry(c2, textvariable=self.reg_name_var, width=24).grid(
            row=3, column=1, sticky="ew", padx=(6, 0), pady=(10, 3))
        tk.Label(c2, text="RTL or C-Spec name (optional)", font=("Segoe UI", 8), fg="#999").grid(
            row=4, column=0, columnspan=2, sticky="w")

        # 3. Mode
        c3 = self._card(parent, "3. Configuration Mode")
        self.mode_var = tk.StringVar(value="Both")
        for val, lbl, desc in [
            ("V",    "V  (LAN SW / Consumer)",  ""),
            ("LM",   "LM  (Non-LAN SW / Corp)", ""),
            ("Both", "Both V and LM",            ""),
        ]:
            ttk.Radiobutton(c3, text=lbl, value=val, variable=self.mode_var).pack(anchor=tk.W, pady=2)

        # 4. Desired value / question
        c4 = self._card(parent, "4. Request / Desired Value")
        self.value_var = tk.StringVar()
        ttk.Entry(c4, textvariable=self.value_var, width=28).pack(fill=tk.X, pady=(0, 4))
        tk.Label(c4, text='"enable WoL"  "0x1"  "explain this bit"',
                 font=("Segoe UI", 8), fg="#999", wraplength=220, justify=tk.LEFT).pack(anchor=tk.W)

        # Buttons
        ttk.Button(parent, text="Ask Claude",
                   command=self._ask).pack(fill=tk.X, pady=(14, 4))
        ttk.Button(parent, text="Clear Chat", command=self._clear_chat).pack(fill=tk.X, pady=2)
        ttk.Button(parent, text="Save Chat",  command=self._save_chat).pack(fill=tk.X, pady=2)
        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        ttk.Button(parent, text="Settings",   command=self._open_settings).pack(fill=tk.X, pady=2)

    def _build_chat(self, parent):
        paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # Conversation history
        hist_frame = ttk.LabelFrame(paned, text="Conversation", padding=4)
        paned.add(hist_frame, weight=4)
        self.chat_display = scrolledtext.ScrolledText(
            hist_frame, wrap=tk.WORD, state=tk.DISABLED,
            font=("Segoe UI", 10), bg="#fafafa", relief=tk.FLAT,
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.tag_configure("user",      foreground="#0055aa", font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_configure("assistant", foreground="#1a1a1a", font=("Segoe UI", 10))
        self.chat_display.tag_configure("system",    foreground="#888888", font=("Segoe UI", 9, "italic"))
        self.chat_display.tag_configure("error",     foreground="#cc0000", font=("Segoe UI", 10))

        # Follow-up
        fu_frame = ttk.LabelFrame(paned, text="Follow-up question  (Ctrl+Enter to send)", padding=4)
        paned.add(fu_frame, weight=1)
        self.input_text = tk.Text(fu_frame, height=4, wrap=tk.WORD,
                                  font=("Segoe UI", 10), relief=tk.SOLID, bd=1)
        self.input_text.pack(fill=tk.BOTH, expand=True)
        self.input_text.bind("<Control-Return>", lambda e: self._send_followup())

        btn_row = ttk.Frame(fu_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="Send  [Ctrl+Enter]",
                   command=self._send_followup).pack(side=tk.RIGHT, padx=4)
        self.status_var = tk.StringVar(value="Fill the form above and click Ask Claude")
        tk.Label(btn_row, textvariable=self.status_var,
                 font=("Segoe UI", 8), fg="#555").pack(side=tk.LEFT, padx=4)

    # ── Projects ──────────────────────────────────────────────────────────

    def _load_projects(self):
        if self.nvm_root and self.nvm_root.exists():
            projects = sorted(d.name for d in self.nvm_root.iterdir() if d.is_dir())
        else:
            projects = [
                "MTL_M_P", "Nahum7_kbl_lp", "Nahum7_spt_lp",
                "Nahum8_cml_lp", "Nahum9_tgl_lp", "Nahum10_adp_lp",
                "Nahum11_lnl_m", "Nahum11_mtl_m_p", "Nahum13_nvl_pch_s",
            ]
        self.project_cb["values"] = projects
        if projects:
            self.project_cb.current(0)
            self._on_project_change()

    def _on_project_change(self):
        proj = self.project_var.get()
        if self.nvm_root:
            p = self.nvm_root / proj
            self.folder_lbl.configure(text=str(p) if p.exists() else "folder not found")
            xlsm_files = list(p.glob("*.xlsm")) if p.exists() else []
            if xlsm_files:
                try:
                    self.excel_context = self._read_excel(xlsm_files[0])
                    self.nvm_status_lbl.configure(
                        text=f"Context loaded: {xlsm_files[0].name}", fg="#107c10")
                except Exception:
                    self.excel_context = ""
            else:
                self.excel_context = ""
        else:
            self.folder_lbl.configure(text=f"({NVM_ROOT} not found — using fallback list)")

    # ── API ───────────────────────────────────────────────────────────────

    def _try_init_client(self):
        if anthropic is None:
            self.api_status_lbl.configure(text="API: anthropic not installed", fg="#cc0000")
            return
        key = self.config_data.get("api_key", "").strip()
        if key:
            self.client = anthropic.Anthropic(api_key=key)
            self.api_status_lbl.configure(text="API: connected", fg="#107c10")
        else:
            self.api_status_lbl.configure(text="API: no key — click Settings", fg="#cc0000")
            self._append_chat("system",
                "No API key configured. Click Settings to add your Anthropic API key.")

    def _build_form_prompt(self):
        project = self.project_var.get()
        offset  = self.offset_var.get().strip()
        bits    = self.bits_var.get().strip()
        name    = self.reg_name_var.get().strip()
        mode    = self.mode_var.get()
        value   = self.value_var.get().strip()

        lines = [f"Project/Platform: {project}"]
        if offset:
            o = offset.upper().lstrip("0X")
            lines.append(f"NVM word offset: 0x{o}")
        if bits:
            lines.append(f"Bit field: [{bits}]")
        if name:
            lines.append(f"Register / field name: {name}")
        lines.append(f"Configuration mode: {mode}")
        if value:
            lines.append(f"Request: {value}")
        else:
            lines.append("Request: Explain this register/bit field and provide the recommended configuration value.")
        return "\n".join(lines)

    def _ask(self):
        if not self._check_client():
            return
        prompt = self._build_form_prompt()
        # Start fresh conversation
        self.conversation.clear()
        self.chat_display.configure(state=tk.NORMAL)
        self.chat_display.delete("1.0", tk.END)
        self.chat_display.configure(state=tk.DISABLED)
        self._submit(prompt)

    def _send_followup(self):
        if not self._check_client():
            return
        msg = self.input_text.get("1.0", tk.END).strip()
        if not msg:
            return
        self.input_text.delete("1.0", tk.END)
        self._submit(msg)

    def _submit(self, user_msg):
        self._append_chat("user", user_msg)
        self.conversation.append({"role": "user", "content": user_msg})
        self.status_var.set("Claude is thinking...")
        self.update_idletasks()
        threading.Thread(target=self._call_api, daemon=True).start()

    def _call_api(self):
        try:
            system = SYSTEM_PROMPT
            if self.excel_context:
                proj = self.project_var.get()
                system += f"\n\n--- NVM MAP CONTEXT ({proj}) ---\n{self.excel_context[:10000]}\n---"
            response = self.client.messages.create(
                model=self.config_data.get("model", MODEL),
                max_tokens=MAX_TOKENS,
                system=system,
                messages=self.conversation,
            )
            reply = response.content[0].text
            self.conversation.append({"role": "assistant", "content": reply})
            self.after(0, lambda: self._append_chat("assistant", reply))
            self.after(0, lambda: self.status_var.set(
                "Done — ask a follow-up or change the form and click Ask Claude again"))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._append_chat("error", f"API Error: {err}"))
            self.after(0, lambda: self.status_var.set("Error — see chat"))

    def _check_client(self):
        if anthropic is None:
            messagebox.showerror("Missing package",
                "Run:  pip install anthropic\nThen restart the app.")
            return False
        if not self.client:
            messagebox.showwarning("No API Key",
                "Configure your Anthropic API key in Settings.")
            self._open_settings()
            return False
        return True

    # ── Chat helpers ──────────────────────────────────────────────────────

    def _append_chat(self, role, text):
        self.chat_display.configure(state=tk.NORMAL)
        labels = {"user": "\nYou:\n", "assistant": "\nClaude:\n",
                  "system": "\n", "error": "\nError: "}
        self.chat_display.insert(tk.END, labels.get(role, "\n"), role)
        self.chat_display.insert(tk.END, text + "\n", role)
        self.chat_display.insert(tk.END, "-" * 60 + "\n", "system")
        self.chat_display.configure(state=tk.DISABLED)
        self.chat_display.see(tk.END)

    def _clear_chat(self):
        if messagebox.askyesno("Clear Chat", "Clear the entire conversation?"):
            self.conversation.clear()
            self.chat_display.configure(state=tk.NORMAL)
            self.chat_display.delete("1.0", tk.END)
            self.chat_display.configure(state=tk.DISABLED)

    def _save_chat(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            lines = [f"[{'You' if m['role'] == 'user' else 'Claude'}]\n{m['content']}\n"
                     for m in self.conversation]
            Path(path).write_text("\n".join(lines), encoding="utf-8")

    # ── Excel ─────────────────────────────────────────────────────────────

    def _read_excel(self, path):
        try:
            from openpyxl import load_workbook
        except ImportError:
            return ""
        wb = load_workbook(str(path), read_only=True, data_only=True, keep_vba=False)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"\n== {ws.title} ==")
            for row in ws.iter_rows(max_row=300, values_only=True):
                if any(c is not None for c in row):
                    lines.append("\t".join("" if c is None else str(c) for c in row))
        wb.close()
        return "\n".join(lines)

    # ── Settings ──────────────────────────────────────────────────────────

    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("500x220")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.columnconfigure(1, weight=1)

        ttk.Label(dlg, text="Anthropic API Key:").grid(row=0, column=0, sticky="w", padx=16, pady=(18, 4))
        key_var = tk.StringVar(value=self.config_data.get("api_key", ""))
        key_entry = ttk.Entry(dlg, textvariable=key_var, width=44, show="*")
        key_entry.grid(row=0, column=1, padx=(4, 16), sticky="ew", pady=(18, 4))

        show_var = tk.BooleanVar()
        ttk.Checkbutton(dlg, text="Show key", variable=show_var,
                        command=lambda: key_entry.configure(show="" if show_var.get() else "*")).grid(
            row=1, column=1, sticky="w", padx=4)

        ttk.Label(dlg, text="Model:").grid(row=2, column=0, sticky="w", padx=16, pady=(12, 4))
        model_var = tk.StringVar(value=self.config_data.get("model", MODEL))
        ttk.Combobox(dlg, textvariable=model_var, width=32,
                     values=["claude-opus-4-5", "claude-sonnet-4-5", "claude-haiku-4-5"]).grid(
            row=2, column=1, sticky="w", padx=(4, 16), pady=(12, 4))

        ttk.Label(dlg, text="Get your key at: console.anthropic.com",
                  foreground="#0071c5", font=("Segoe UI", 8)).grid(
            row=3, column=0, columnspan=2, sticky="w", padx=16, pady=(4, 0))

        def _save():
            self.config_data["api_key"] = key_var.get().strip()
            self.config_data["model"]   = model_var.get()
            save_config(self.config_data)
            self._try_init_client()
            dlg.destroy()

        btns = ttk.Frame(dlg)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", padx=16, pady=16)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="Save",   command=_save).pack(side=tk.RIGHT)


def main():
    app = NvmAiApp()
    app.mainloop()


if __name__ == "__main__":
    main()
