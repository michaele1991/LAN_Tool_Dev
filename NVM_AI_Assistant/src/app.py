"""
NVM AI Assistant — Intel GBE NVM Configuration Advisor
Powered by Anthropic Claude API

Generates NVM configuration suggestions based on user requirements.
Can read the current project's Excel data as context.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

try:
    import anthropic
except ImportError:
    anthropic = None

APP_TITLE   = "Intel® GBE NVM AI Assistant"
APP_VERSION = "v1.0"
MODEL       = "claude-opus-4-5"
MAX_TOKENS  = 4096
CONFIG_FILE = Path(__file__).parent / "config.json"

SYSTEM_PROMPT = """You are an expert Intel GBE (Gigabit Ethernet) NVM (Non-Volatile Memory) configuration engineer.
You help engineers generate, review, and optimize NVM configurations for Intel Ethernet controllers
(Nahum7, Nahum8, Nahum9, Nahum10, Nahum11, Nahum13, MTL and similar platforms).

Your expertise includes:
- GBE NVM word/bit field definitions and their effects
- LAN SW vs Non-LAN SW (LM/V) configuration differences  
- Device ID, SKU, silicon stepping considerations
- Wake-on-LAN, PCIe, power management settings
- Checksum calculation (word 0x3F = 0xBABA)
- C-Spec and RTL register naming conventions

When asked to generate a configuration:
1. List the relevant word indices (hex offset) and bit fields to modify
2. Specify the exact value (V and/or LM) with rationale
3. Flag any dependencies or ordering requirements
4. Mention any risks or platform-specific caveats

Format register changes as a table when possible:
| Offset | Bits | Field | Value | Reason |
"""


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"api_key": "", "last_project": ""}


def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


class NvmAiApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  {APP_VERSION}")
        self.geometry("1100x780")
        self.minsize(800, 600)
        self.config_data = load_config()
        self.client = None
        self.conversation: list[dict] = []
        self.context_text = ""
        self._build_ui()
        self._try_init_client()

    # ── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._apply_styles()

        # Header
        hdr = tk.Frame(self, bg="#0071c5", height=52)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  {APP_TITLE}  {APP_VERSION}",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#0071c5").pack(side=tk.LEFT, padx=12, pady=10)
        tk.Label(hdr, text="Powered by Claude (Anthropic)",
                 font=("Segoe UI", 9), fg="#a8d4f0", bg="#0071c5").pack(side=tk.RIGHT, padx=16)

        # Toolbar
        toolbar = ttk.Frame(self, padding=(8, 4))
        toolbar.pack(fill=tk.X)
        ttk.Button(toolbar, text="⚙  Settings",        command=self._open_settings).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="📂  Load NVM Context", command=self._load_context).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="🗑  Clear Chat",        command=self._clear_chat).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="💾  Save Chat",         command=self._save_chat).pack(side=tk.LEFT, padx=4)
        self.ctx_label = tk.Label(toolbar, text="No context loaded", fg="#888",
                                  font=("Segoe UI", 8))
        self.ctx_label.pack(side=tk.RIGHT, padx=8)

        # Main pane
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 0))

        # Chat history
        chat_frame = ttk.LabelFrame(paned, text="Conversation", padding=(4, 4))
        paned.add(chat_frame, weight=4)

        self.chat_display = scrolledtext.ScrolledText(
            chat_frame, wrap=tk.WORD, state=tk.DISABLED,
            font=("Segoe UI", 10), bg="#fafafa", relief=tk.FLAT
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.tag_configure("user",      foreground="#0055aa", font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_configure("assistant", foreground="#1a1a1a", font=("Segoe UI", 10))
        self.chat_display.tag_configure("system",    foreground="#888888", font=("Segoe UI", 9, "italic"))
        self.chat_display.tag_configure("error",     foreground="#cc0000", font=("Segoe UI", 9, "italic"))

        # Input area
        input_frame = ttk.LabelFrame(paned, text="Your Request", padding=(4, 4))
        paned.add(input_frame, weight=1)

        self.input_text = tk.Text(input_frame, height=5, wrap=tk.WORD,
                                  font=("Segoe UI", 10), relief=tk.SOLID, bd=1)
        self.input_text.pack(fill=tk.BOTH, expand=True)
        self.input_text.bind("<Control-Return>", lambda e: self._send())

        btn_row = ttk.Frame(input_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="Send  [Ctrl+Enter]", command=self._send,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=4)
        self.status_var = tk.StringVar(value="Ready — type a request and press Send")
        tk.Label(btn_row, textvariable=self.status_var, fg="#555",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=4)

        # Quick prompts
        qp_frame = ttk.LabelFrame(self, text="Quick Prompts", padding=(4, 2))
        qp_frame.pack(fill=tk.X, padx=8, pady=(2, 0))
        prompts = [
            ("Wake-on-LAN enable",     "Generate NVM configuration to enable Wake-on-LAN (magic packet) on all power states."),
            ("PCIe power down disable","Generate config to prevent PCIe link from powering down in D3 state."),
            ("LM vs V diff",           "Explain the key differences between LM and V NVM configurations for this platform."),
            ("Checksum verify",        "How do I verify the NVM checksum is correct? What is the algorithm?"),
            ("Device ID set",          "What words/bits control the Device ID and SKU Device ID in the NVM map?"),
        ]
        for label, prompt in prompts:
            ttk.Button(qp_frame, text=label, width=22,
                       command=lambda p=prompt: self._set_input(p)).pack(side=tk.LEFT, padx=2, pady=2)

        # Status bar
        sb = tk.Frame(self, bg="#e8e8e8", height=22, relief=tk.SUNKEN, bd=1)
        sb.pack(side=tk.BOTTOM, fill=tk.X)
        sb.pack_propagate(False)
        self.api_status = tk.Label(sb, text="API: not connected", fg="#888",
                                   font=("Segoe UI", 8), bg="#e8e8e8")
        self.api_status.pack(side=tk.LEFT, padx=8)
        tk.Label(sb, text=f"{APP_TITLE}  {APP_VERSION}",
                 font=("Segoe UI", 8), bg="#e8e8e8", fg="#aaa").pack(side=tk.RIGHT, padx=8)

    def _apply_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("vista")
        except Exception:
            pass
        s.configure("Accent.TButton", foreground="white", background="#0071c5",
                     font=("Segoe UI", 9, "bold"))

    # ── Client ──────────────────────────────────────────────────────────────

    def _try_init_client(self):
        if anthropic is None:
            self.api_status.configure(text="API: anthropic package not installed", fg="#cc0000")
            return
        key = self.config_data.get("api_key", "").strip()
        if key:
            self.client = anthropic.Anthropic(api_key=key)
            self.api_status.configure(text="API: connected ✓", fg="#107c10")
        else:
            self.api_status.configure(text="API: no key — click Settings", fg="#cc0000")
            self._append_chat("system", "No API key configured. Click ⚙ Settings to add your Anthropic API key.")

    # ── Actions ─────────────────────────────────────────────────────────────

    def _send(self):
        if anthropic is None:
            messagebox.showerror("Missing package",
                                 "anthropic package is not installed.\nRun: pip install anthropic")
            return
        if not self.client:
            messagebox.showwarning("No API Key", "Configure your Anthropic API key in Settings first.")
            self._open_settings()
            return

        user_msg = self.input_text.get("1.0", tk.END).strip()
        if not user_msg:
            return

        self.input_text.delete("1.0", tk.END)
        self._append_chat("user", user_msg)
        self.conversation.append({"role": "user", "content": user_msg})
        self.status_var.set("Thinking…")
        self.update_idletasks()

        import threading
        threading.Thread(target=self._call_api, daemon=True).start()

    def _call_api(self):
        try:
            # Build system prompt — optionally inject NVM context
            system = SYSTEM_PROMPT
            if self.context_text:
                system += f"\n\n--- LOADED NVM CONTEXT ---\n{self.context_text[:8000]}\n---"

            response = self.client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=system,
                messages=self.conversation,
            )
            reply = response.content[0].text
            self.conversation.append({"role": "assistant", "content": reply})
            self.after(0, lambda: self._append_chat("assistant", reply))
            self.after(0, lambda: self.status_var.set("Ready"))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._append_chat("error", f"Error: {err}"))
            self.after(0, lambda: self.status_var.set("Error — see chat"))

    def _append_chat(self, role: str, text: str):
        self.chat_display.configure(state=tk.NORMAL)
        prefix = {"user": "\n🧑 You:\n", "assistant": "\n🤖 Claude:\n",
                  "system": "\n", "error": "\n⚠ "}.get(role, "\n")
        self.chat_display.insert(tk.END, prefix, role)
        self.chat_display.insert(tk.END, text + "\n", role)
        self.chat_display.insert(tk.END, "─" * 60 + "\n", "system")
        self.chat_display.configure(state=tk.DISABLED)
        self.chat_display.see(tk.END)

    def _set_input(self, text: str):
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert("1.0", text)
        self.input_text.focus_set()

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
            title="Save Chat"
        )
        if path:
            lines = []
            for msg in self.conversation:
                role = "You" if msg["role"] == "user" else "Claude"
                lines.append(f"[{role}]\n{msg['content']}\n")
            Path(path).write_text("\n".join(lines), encoding="utf-8")

    def _load_context(self):
        """Load a .xlsx/.xlsm data file or a .json/txt as context for the AI."""
        path = filedialog.askopenfilename(
            title="Load NVM Context File",
            filetypes=[
                ("Excel files", "*.xlsm *.xlsx"),
                ("JSON files",  "*.json"),
                ("Text files",  "*.txt"),
                ("All files",   "*.*"),
            ]
        )
        if not path:
            return
        p = Path(path)
        try:
            if p.suffix.lower() in (".json",):
                data = json.loads(p.read_text(encoding="utf-8"))
                self.context_text = json.dumps(data, indent=2)
            elif p.suffix.lower() in (".txt",):
                self.context_text = p.read_text(encoding="utf-8", errors="replace")
            elif p.suffix.lower() in (".xlsm", ".xlsx"):
                self.context_text = self._read_excel_context(p)
            self.ctx_label.configure(text=f"Context: {p.name}  ({len(self.context_text)} chars)", fg="#107c10")
            self._append_chat("system", f"Context loaded: {p.name}")
        except Exception as e:
            messagebox.showerror("Load Context", f"Failed to load context:\n{e}")

    def _read_excel_context(self, path: Path) -> str:
        try:
            from openpyxl import load_workbook
        except ImportError:
            return f"[openpyxl not installed — cannot read {path.name}]"
        wb = load_workbook(str(path), read_only=True, data_only=True, keep_vba=False)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"\n== Sheet: {ws.title} ==")
            for row in ws.iter_rows(max_row=200, values_only=True):
                if any(c is not None for c in row):
                    lines.append("\t".join(str(c) if c is not None else "" for c in row))
        wb.close()
        return "\n".join(lines)

    # ── Settings dialog ─────────────────────────────────────────────────────

    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("480x220")
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text="Anthropic API Key:", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=16, pady=(20, 4))
        key_var = tk.StringVar(value=self.config_data.get("api_key", ""))
        key_entry = ttk.Entry(dlg, textvariable=key_var, width=44, show="*")
        key_entry.grid(row=1, column=0, columnspan=2, padx=16, sticky="ew")

        show_var = tk.BooleanVar()
        def _toggle_show():
            key_entry.configure(show="" if show_var.get() else "*")
        ttk.Checkbutton(dlg, text="Show key", variable=show_var,
                        command=_toggle_show).grid(row=2, column=0, sticky="w", padx=16, pady=4)

        ttk.Label(dlg, text="Model:", font=("Segoe UI", 9)).grid(
            row=3, column=0, sticky="w", padx=16, pady=(8, 2))
        model_var = tk.StringVar(value=MODEL)
        ttk.Combobox(dlg, textvariable=model_var, width=30,
                     values=["claude-opus-4-5", "claude-sonnet-4-5", "claude-haiku-4-5"]).grid(
            row=4, column=0, padx=16, sticky="w")

        def _save():
            self.config_data["api_key"] = key_var.get().strip()
            global MODEL
            MODEL = model_var.get()
            save_config(self.config_data)
            self._try_init_client()
            dlg.destroy()

        btn_f = ttk.Frame(dlg)
        btn_f.grid(row=5, column=0, columnspan=2, sticky="e", padx=16, pady=12)
        ttk.Button(btn_f, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_f, text="Save", style="Accent.TButton", command=_save).pack(side=tk.RIGHT)

        ttk.Label(dlg, text="Get your key at: console.anthropic.com",
                  foreground="#0071c5", font=("Segoe UI", 8)).grid(
            row=6, column=0, columnspan=2, padx=16, sticky="w")


def main():
    if anthropic is None:
        import subprocess
        answer = input("anthropic package not found. Install now? [y/N]: ")
        if answer.strip().lower() == "y":
            subprocess.run([sys.executable, "-m", "pip", "install", "anthropic"], check=True)
        else:
            sys.exit(1)
    app = NvmAiApp()
    app.mainloop()


if __name__ == "__main__":
    main()
