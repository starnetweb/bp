"""
Academic Research Writeup Agent — Desktop GUI
A Tkinter interface for research_agent.py
Run this file to open the application window.
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog

# ── make sure the agent module is importable from same folder ──
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import research_agent


# ─────────────────────────────────────────────────────────
#  COLOUR / STYLE CONSTANTS
# ─────────────────────────────────────────────────────────
BG_DARK    = "#1E2A3A"
BG_MID     = "#243447"
BG_CARD    = "#2C3E52"
ACCENT     = "#4A90D9"
ACCENT_HOV = "#5BA3F5"
TEXT_WHITE = "#F0F4F8"
TEXT_GREY  = "#8FA8C8"
SUCCESS    = "#4CAF50"
WARNING    = "#FFA726"
ERROR      = "#EF5350"
BORDER     = "#3D5470"


# ─────────────────────────────────────────────────────────
#  MAIN APP WINDOW
# ─────────────────────────────────────────────────────────
class ResearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Academic Research Writeup Agent")
        self.geometry("820x720")
        self.minsize(700, 600)
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        # State
        self._running   = False
        self._out_path  = None
        self._api_key_var   = tk.StringVar()
        self._topic_var     = tk.StringVar()
        self._save_dir_var  = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        self._show_key_var  = tk.BooleanVar(value=False)

        # Pre-fill API key from environment if available
        env_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if env_key:
            self._api_key_var.set(env_key)

        self._build_ui()

    # ── UI CONSTRUCTION ──────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_form()
        self._build_log()
        self._build_statusbar()

    def _build_header(self):
        hdr = tk.Frame(self, bg=BG_MID, pady=14)
        hdr.pack(fill="x")

        tk.Label(
            hdr, text="📄  Academic Research Writeup Agent",
            font=("Segoe UI", 17, "bold"),
            bg=BG_MID, fg=TEXT_WHITE
        ).pack()

        tk.Label(
            hdr, text="5-Chapter Academic Document  ·  Powered by Claude Opus 4.6",
            font=("Segoe UI", 10),
            bg=BG_MID, fg=TEXT_GREY
        ).pack(pady=(2, 0))

    def _build_form(self):
        card = tk.Frame(self, bg=BG_CARD, padx=24, pady=18,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=20, pady=(14, 0))

        # ── API Key row ──
        tk.Label(card, text="Anthropic API Key",
                 font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_GREY).grid(row=0, column=0, sticky="w", pady=(0, 4))

        key_frame = tk.Frame(card, bg=BG_CARD)
        key_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        card.columnconfigure(0, weight=1)

        self._key_entry = tk.Entry(
            key_frame, textvariable=self._api_key_var,
            show="•", font=("Segoe UI", 11),
            bg=BG_MID, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor=ACCENT, highlightbackground=BORDER
        )
        self._key_entry.pack(side="left", fill="x", expand=True, ipady=8, ipadx=10)

        toggle_btn = tk.Button(
            key_frame, text="Show",
            font=("Segoe UI", 9),
            bg=BG_MID, fg=TEXT_GREY,
            relief="flat", bd=0, cursor="hand2",
            command=self._toggle_key_visibility
        )
        toggle_btn.pack(side="left", padx=(6, 0))
        self._toggle_btn = toggle_btn

        # ── Topic row ──
        tk.Label(card, text="Research Topic",
                 font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_GREY).grid(row=2, column=0, sticky="w",
                                                 pady=(14, 4))

        self._topic_entry = tk.Entry(
            card, textvariable=self._topic_var,
            font=("Segoe UI", 12),
            bg=BG_MID, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor=ACCENT, highlightbackground=BORDER
        )
        self._topic_entry.grid(row=3, column=0, columnspan=2,
                               sticky="ew", ipady=10, ipadx=10)
        self._topic_entry.bind("<Return>", lambda e: self._start())

        placeholder = "e.g.  The Impact of Artificial Intelligence on Healthcare in Africa"
        self._topic_entry.insert(0, placeholder)
        self._topic_entry.config(fg=TEXT_GREY)
        self._topic_entry.bind("<FocusIn>",  self._clear_placeholder)
        self._topic_entry.bind("<FocusOut>", self._restore_placeholder)
        self._placeholder = placeholder

        # ── Save folder row ──
        tk.Label(card, text="Save Folder",
                 font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=TEXT_GREY).grid(row=4, column=0, sticky="w",
                                                  pady=(14, 4))

        dir_frame = tk.Frame(card, bg=BG_CARD)
        dir_frame.grid(row=5, column=0, columnspan=2, sticky="ew")

        self._dir_entry = tk.Entry(
            dir_frame, textvariable=self._save_dir_var,
            font=("Segoe UI", 11),
            bg=BG_MID, fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor=ACCENT, highlightbackground=BORDER
        )
        self._dir_entry.pack(side="left", fill="x", expand=True, ipady=8, ipadx=10)

        browse_btn = tk.Button(
            dir_frame, text="Browse",
            font=("Segoe UI", 9),
            bg=ACCENT, fg=TEXT_WHITE,
            activebackground=ACCENT_HOV, activeforeground=TEXT_WHITE,
            relief="flat", bd=0, cursor="hand2", padx=12, pady=4,
            command=self._browse_folder
        )
        browse_btn.pack(side="left", padx=(6, 0))

        # ── Generate button ──
        self._gen_btn = tk.Button(
            card, text="⚡  Generate Research Document",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT, fg=TEXT_WHITE,
            activebackground=ACCENT_HOV, activeforeground=TEXT_WHITE,
            relief="flat", bd=0, cursor="hand2", pady=12,
            command=self._start
        )
        self._gen_btn.grid(row=6, column=0, columnspan=2,
                           sticky="ew", pady=(20, 4))

        # ── Open file button (hidden until done) ──
        self._open_btn = tk.Button(
            card, text="📂  Open Document",
            font=("Segoe UI", 11, "bold"),
            bg=SUCCESS, fg=TEXT_WHITE,
            activebackground="#43A047", activeforeground=TEXT_WHITE,
            relief="flat", bd=0, cursor="hand2", pady=10,
            command=self._open_document
        )
        self._open_btn.grid(row=7, column=0, columnspan=2,
                            sticky="ew", pady=(0, 4))
        self._open_btn.grid_remove()   # hidden initially

    def _build_log(self):
        log_frame = tk.Frame(self, bg=BG_DARK, padx=20, pady=10)
        log_frame.pack(fill="both", expand=True, pady=(12, 0))

        header = tk.Frame(log_frame, bg=BG_DARK)
        header.pack(fill="x")

        tk.Label(header, text="Progress Log",
                 font=("Segoe UI", 10, "bold"),
                 bg=BG_DARK, fg=TEXT_GREY).pack(side="left")

        tk.Button(
            header, text="Clear",
            font=("Segoe UI", 8),
            bg=BG_MID, fg=TEXT_GREY,
            relief="flat", bd=0, cursor="hand2",
            command=self._clear_log
        ).pack(side="right")

        self._log = scrolledtext.ScrolledText(
            log_frame,
            font=("Consolas", 10),
            bg="#111D2B", fg=TEXT_WHITE,
            insertbackground=TEXT_WHITE,
            relief="flat", bd=0,
            state="disabled",
            wrap="word",
            padx=10, pady=8
        )
        self._log.pack(fill="both", expand=True, pady=(6, 0))

        # Colour tags for log
        self._log.tag_config("info",    foreground=TEXT_WHITE)
        self._log.tag_config("success", foreground=SUCCESS)
        self._log.tag_config("warning", foreground=WARNING)
        self._log.tag_config("error",   foreground=ERROR)
        self._log.tag_config("accent",  foreground=ACCENT)
        self._log.tag_config("header",  foreground=ACCENT, font=("Consolas", 10, "bold"))

        # Progress bar
        self._progress = ttk.Progressbar(log_frame, mode="indeterminate", length=400)
        self._progress.pack(fill="x", pady=(8, 0))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TProgressbar",
                         troughcolor=BG_MID,
                         background=ACCENT,
                         thickness=6)

    def _build_statusbar(self):
        self._status_var = tk.StringVar(value="Ready — enter a topic and click Generate.")
        bar = tk.Label(
            self, textvariable=self._status_var,
            font=("Segoe UI", 9), anchor="w",
            bg=BG_MID, fg=TEXT_GREY, padx=14, pady=5
        )
        bar.pack(fill="x", side="bottom")

    # ── ACTIONS ─────────────────────────────────────────

    def _toggle_key_visibility(self):
        if self._key_entry.cget("show") == "•":
            self._key_entry.config(show="")
            self._toggle_btn.config(text="Hide")
        else:
            self._key_entry.config(show="•")
            self._toggle_btn.config(text="Show")

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select Save Folder")
        if folder:
            self._save_dir_var.set(folder)

    def _clear_placeholder(self, event):
        if self._topic_var.get() == self._placeholder:
            self._topic_entry.delete(0, "end")
            self._topic_entry.config(fg=TEXT_WHITE)

    def _restore_placeholder(self, event):
        if not self._topic_var.get().strip():
            self._topic_entry.insert(0, self._placeholder)
            self._topic_entry.config(fg=TEXT_GREY)

    def _start(self):
        if self._running:
            return

        topic = self._topic_var.get().strip()
        if not topic or topic == self._placeholder:
            messagebox.showwarning("Missing Topic", "Please enter a research topic.")
            return

        api_key = self._api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("Missing API Key",
                                   "Please enter your Anthropic API key.")
            return

        save_dir = self._save_dir_var.get().strip() or os.getcwd()
        if not os.path.isdir(save_dir):
            messagebox.showerror("Invalid Folder",
                                 f"The save folder does not exist:\n{save_dir}")
            return

        os.environ["ANTHROPIC_API_KEY"] = api_key
        self._running  = True
        self._out_path = None
        self._open_btn.grid_remove()

        self._gen_btn.config(state="disabled",
                             text="⏳  Generating — please wait...",
                             bg="#3A6090")
        self._progress.start(12)
        self._clear_log()
        self._set_status("Generating... this takes 2–4 minutes.")

        # Run in background thread so the UI stays responsive
        thread = threading.Thread(
            target=self._run_agent,
            args=(topic, save_dir),
            daemon=True
        )
        thread.start()

    def _run_agent(self, topic, save_dir):
        """Background worker — calls the agent and streams log updates."""
        import io, contextlib

        def log(msg, tag="info"):
            self.after(0, self._append_log, msg, tag)

        try:
            log("=" * 56, "header")
            log(f"  TOPIC: {topic}", "header")
            log("=" * 56, "header")
            log("")

            client = research_agent.anthropic.Anthropic()

            # Front matter
            log("► Generating front matter (Abstract / Acknowledgements)...", "accent")
            front = research_agent.generate_front_matter(client, topic)
            log(f"  ✓ Front matter — {len(front):,} characters\n", "success")

            # 5 chapters
            chapters = {}
            for num in range(1, 6):
                ch_name = research_agent.CHAPTER_SUBTITLES[num]
                log(f"► Chapter {num}: {ch_name}...", "accent")
                chapters[num] = research_agent.generate_chapter(client, topic, num)
                log(f"  ✓ Chapter {num} done — {len(chapters[num]):,} characters\n",
                    "success")

            # Build document
            log("► Assembling Word document...", "accent")
            import re, os
            safe = re.sub(r"[^\w\s-]", "", topic).strip().replace(" ", "_")[:50]
            filename = f"Research_{safe}.docx"
            out_path = os.path.join(save_dir, filename)

            # Patch save path inside build_document
            doc = research_agent.Document()
            research_agent.set_document_defaults(doc)
            research_agent.build_title_page(doc, topic)
            research_agent.build_front_matter_page(doc, front)
            research_agent.build_toc_page(doc, topic)
            research_agent.build_abbreviations_page(doc)
            for num in range(1, 6):
                research_agent.build_chapter_page(doc, num, chapters[num])
            doc.save(out_path)

            self._out_path = out_path

            log("")
            log("=" * 56, "header")
            log("  ✅  DOCUMENT READY!", "success")
            log(f"  📄  {out_path}", "success")
            log("=" * 56, "header")

            self.after(0, self._on_success)

        except Exception as exc:
            log(f"\n❌ Error: {exc}", "error")
            import traceback
            log(traceback.format_exc(), "error")
            self.after(0, self._on_error, str(exc))

    def _on_success(self):
        self._running = False
        self._progress.stop()
        self._progress.config(value=100)
        self._gen_btn.config(state="normal",
                             text="⚡  Generate Research Document",
                             bg=ACCENT)
        self._open_btn.grid()
        self._set_status(f"Done!  Saved to: {self._out_path}")

    def _on_error(self, msg):
        self._running = False
        self._progress.stop()
        self._gen_btn.config(state="normal",
                             text="⚡  Generate Research Document",
                             bg=ACCENT)
        self._set_status(f"Error: {msg}")
        messagebox.showerror("Generation Failed", f"An error occurred:\n\n{msg}")

    def _open_document(self):
        if self._out_path and os.path.exists(self._out_path):
            if sys.platform == "win32":
                os.startfile(self._out_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self._out_path])
            else:
                subprocess.Popen(["xdg-open", self._out_path])
        else:
            messagebox.showerror("File Not Found",
                                 "Could not locate the generated file.")

    # ── LOG HELPERS ──────────────────────────────────────

    def _append_log(self, msg, tag="info"):
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _set_status(self, msg):
        self._status_var.set(msg)


# ─────────────────────────────────────────────────────────
#  LAUNCH
# ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ResearchApp()
    app.mainloop()
