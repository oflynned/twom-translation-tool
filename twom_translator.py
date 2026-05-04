"""
TWoM Translation Tool  v5.0

Workflow:
  1. Import .lang  →  creates a working CSV
  2. Edit translations in this tool (or open CSV in Excel / share with team)
  3. Pack to .lang  →  produces binary file for Storyteller
  4. Open in Storyteller → Pack → test in game
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, json, threading, urllib.request, glob, csv
from pathlib import Path
from lang_parser import load_lang, build_lang, load_csv, save_csv, LangRecord

CFG = Path.home() / ".twom_v5.json"
def load_cfg():
    try: return json.loads(CFG.read_text()) if CFG.exists() else {}
    except: return {}
def save_cfg(c):
    try: CFG.write_text(json.dumps(c, indent=2))
    except: pass

def ai_translate(text, src, tgt, key):
    body = json.dumps({
        "model": "claude-sonnet-4-20250514", "max_tokens": 512,
        "messages": [{"role": "user", "content":
            f"Translate this {src} game text to {tgt} for 'This War of Mine' "
            f"(somber survival game). Keep {{0}}, %s, \\n intact. "
            f"Return only the translated text.\n\n{text}"}]
    }).encode()
    import json as _json
    req = urllib.request.Request("https://api.anthropic.com/v1/messages",
                                 data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("x-api-key", key)
    req.add_header("anthropic-version", "2023-06-01")
    with urllib.request.urlopen(req, timeout=30) as r:
        return _json.loads(r.read())["content"][0]["text"].strip()


# ── Palette ────────────────────────────────────────────────────────────────
BG      = "#FAFAF9"      # near-white warm
BG2     = "#F2F1EF"      # sidebar / panel
BORDER  = "#E2E1DE"      # dividers
TEXT    = "#1C1C1A"      # primary text
TEXT2   = "#6B6B68"      # secondary
ACCENT  = "#1D6F5A"      # teal-green action color
ACCENTL = "#EAF4F1"      # accent tint (selected row bg)
ACCENTT = "#FFFFFF"      # text on accent button
DANGER  = "#C0392B"      # destructive / warning
DONE    = "#27AE60"      # translated indicator

FONT    = ("Segoe UI", 10)       # Windows; falls back gracefully on other OS
FONTB   = ("Segoe UI", 10, "bold")
FONTS   = ("Segoe UI", 9)
FONTM   = ("Consolas", 9)        # monospace for keys


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TWoM Translation Tool")
        self.geometry("1280x800")
        self.minsize(900, 580)
        self.configure(bg=BG)

        self.cfg        = load_cfg()
        self.api_key    = tk.StringVar(value=self.cfg.get("api_key", ""))
        self.tgt_lang   = tk.StringVar(value=self.cfg.get("tgt_lang", "Irish"))
        self.csv_path   : Path | None       = None
        self.records    : list[LangRecord]  = []   # key + source value
        self.translations: dict[str, str]   = {}   # key → translated text
        self.changed    = False
        self.ai_busy    = False
        self._cur_grp   : str | None        = None
        self._grp_ids   : dict              = {}   # treeview iid → group str|None
        self._tv_items  : list[str]         = []   # ordered treeview iids
        self._search_q  = tk.StringVar()
        self._show_empty = tk.BooleanVar(value=False)
        self._status    = tk.StringVar(value="Open a .lang file or existing CSV to begin.")
        self._progress  = tk.StringVar(value="")

        self._build_styles()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Styles ────────────────────────────────────────────────────────────
    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")

        # Base
        s.configure(".", background=BG, foreground=TEXT,
                    fieldbackground=BG, font=FONT, borderwidth=0)
        s.configure("TFrame", background=BG)

        # Sidebar
        s.configure("Side.TFrame", background=BG2)

        # Labels
        s.configure("TLabel",    background=BG,  foreground=TEXT,  font=FONT)
        s.configure("Dim.TLabel",background=BG,  foreground=TEXT2, font=FONTS)
        s.configure("Side.TLabel",background=BG2,foreground=TEXT2, font=FONTS)

        # Entries
        s.configure("TEntry", fieldbackground="#FFFFFF", foreground=TEXT,
                    insertcolor=ACCENT, borderwidth=1, relief="flat",
                    padding=(6, 5))
        s.map("TEntry", fieldbackground=[("focus", "#FFFFFF")])

        # Primary button
        s.configure("Primary.TButton",
            background=ACCENT, foreground=ACCENTT, font=FONTB,
            borderwidth=0, relief="flat", padding=(14, 7))
        s.map("Primary.TButton",
            background=[("active", "#165947"), ("pressed", "#0F3D31")])

        # Ghost button (toolbar actions)
        s.configure("Ghost.TButton",
            background=BG, foreground=TEXT, font=FONT,
            borderwidth=1, relief="flat", padding=(10, 5))
        s.map("Ghost.TButton",
            background=[("active", BG2)],
            foreground=[("active", ACCENT)])

        # Icon-only button
        s.configure("Icon.TButton",
            background=BG2, foreground=ACCENT, font=("Segoe UI", 11),
            borderwidth=0, relief="flat", padding=(4, 2))
        s.map("Icon.TButton",
            background=[("active", BORDER)])

        # Treeview (string table)
        s.configure("Strings.Treeview",
            background="#FFFFFF", foreground=TEXT,
            fieldbackground="#FFFFFF", rowheight=28,
            font=FONT, borderwidth=0)
        s.configure("Strings.Treeview.Heading",
            background=BG2, foreground=TEXT2, font=FONTS,
            relief="flat", borderwidth=0, padding=(8, 6))
        s.map("Strings.Treeview",
            background=[("selected", ACCENTL)],
            foreground=[("selected", ACCENT)])

        # Group treeview (sidebar)
        s.configure("Groups.Treeview",
            background=BG2, foreground=TEXT,
            fieldbackground=BG2, rowheight=26,
            font=FONTS, borderwidth=0)
        s.configure("Groups.Treeview.Heading",
            background=BG2, foreground=TEXT2, font=FONTS, relief="flat")
        s.map("Groups.Treeview",
            background=[("selected", ACCENTL)],
            foreground=[("selected", ACCENT)])

        # Scrollbar
        s.configure("TScrollbar", background=BG2, troughcolor=BG,
                    arrowcolor=TEXT2, borderwidth=0, width=8)
        s.map("TScrollbar", background=[("active", BORDER)])

        # Progressbar
        s.configure("TProgressbar", background=ACCENT, troughcolor=BG2,
                    borderwidth=0, thickness=3)

        # Checkbutton
        s.configure("TCheckbutton", background=BG, foreground=TEXT2, font=FONTS)
        s.map("TCheckbutton", background=[("active", BG)], foreground=[("active", TEXT)])

        # Separator
        s.configure("TSeparator", background=BORDER)

    # ── UI ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Top bar ─────────────────────────────────────────────────────
        topbar = tk.Frame(self, bg=BG, pady=0)
        topbar.pack(fill="x")

        # Left section: file actions
        left_bar = tk.Frame(topbar, bg=BG, padx=16, pady=10)
        left_bar.pack(side="left")

        tk.Label(left_bar, text="TWoM Translation",
                 bg=BG, fg=TEXT, font=("Segoe UI", 12, "bold")).pack(side="left", padx=(0,20))

        ttk.Button(left_bar, text="Import .lang",
                   command=self._import_lang, style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(left_bar, text="Open CSV",
                   command=self._open_csv, style="Ghost.TButton").pack(side="left", padx=2)
        ttk.Button(left_bar, text="Save CSV",
                   command=self._save_csv_action, style="Ghost.TButton").pack(side="left", padx=2)

        tk.Frame(left_bar, bg=BORDER, width=1).pack(side="left", fill="y", padx=10, pady=4)

        ttk.Button(left_bar, text="⬇  Pack to .lang",
                   command=self._pack_lang, style="Primary.TButton").pack(side="left", padx=2)

        # Right section: language + status
        right_bar = tk.Frame(topbar, bg=BG, padx=16, pady=10)
        right_bar.pack(side="right")

        tk.Label(right_bar, text="Target language",
                 bg=BG, fg=TEXT2, font=FONTS).pack(side="left", padx=(0,6))
        lang_entry = ttk.Entry(right_bar, textvariable=self.tgt_lang, width=14)
        lang_entry.pack(side="left", padx=(0,16))

        self._prog_bar = ttk.Progressbar(right_bar, orient="horizontal",
                                          mode="determinate", length=100)
        self._prog_bar.pack(side="left", padx=(0,6))
        tk.Label(right_bar, textvariable=self._progress,
                 bg=BG, fg=TEXT2, font=FONTS).pack(side="left")

        # Divider
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Search bar ──────────────────────────────────────────────────
        searchbar = tk.Frame(self, bg=BG2, padx=16, pady=7)
        searchbar.pack(fill="x")

        tk.Label(searchbar, text="Search", bg=BG2, fg=TEXT2, font=FONTS).pack(side="left", padx=(0,6))
        ttk.Entry(searchbar, textvariable=self._search_q, width=32).pack(side="left")
        self._search_q.trace_add("write", lambda *_: self._refresh_table())

        ttk.Checkbutton(searchbar, text="Untranslated only",
                        variable=self._show_empty,
                        command=self._refresh_table).pack(side="left", padx=16)

        self._file_label = tk.Label(searchbar, text="No file", bg=BG2, fg=TEXT2, font=FONTS)
        self._file_label.pack(side="right")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # ── Main content: sidebar + table ───────────────────────────────
        content = tk.Frame(self, bg=BG)
        content.pack(fill="both", expand=True)

        # Sidebar
        sidebar = tk.Frame(content, bg=BG2, width=200)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="GROUPS", bg=BG2, fg=TEXT2,
                 font=("Segoe UI", 8), pady=10, padx=12).pack(anchor="w")

        self.grp_tv = ttk.Treeview(sidebar, show="tree",
                                    selectmode="browse", style="Groups.Treeview")
        grp_sb = ttk.Scrollbar(sidebar, orient="vertical", command=self.grp_tv.yview)
        self.grp_tv.configure(yscrollcommand=grp_sb.set)
        grp_sb.pack(side="right", fill="y")
        self.grp_tv.pack(fill="both", expand=True)
        self.grp_tv.bind("<<TreeviewSelect>>", self._on_group_select)

        tk.Frame(content, bg=BORDER, width=1).pack(side="left", fill="y")

        # Table + edit panel (right side, stacked vertically)
        right_side = tk.Frame(content, bg=BG)
        right_side.pack(side="left", fill="both", expand=True)

        # String table (Treeview — handles 13k rows natively, no hanging)
        table_frame = tk.Frame(right_side, bg=BG)
        table_frame.pack(fill="both", expand=True)

        self.table = ttk.Treeview(
            table_frame,
            columns=("source", "translation", "status"),
            show="headings",
            selectmode="browse",
            style="Strings.Treeview"
        )
        self.table.heading("source",      text="Source")
        self.table.heading("translation", text="Translation")
        self.table.heading("status",      text="")

        self.table.column("source",      width=380, minwidth=180, stretch=True)
        self.table.column("translation", width=380, minwidth=180, stretch=True)
        self.table.column("status",      width=26,  minwidth=26,  stretch=False)

        self.table.tag_configure("done",  foreground=DONE)
        self.table.tag_configure("todo",  foreground=TEXT2)
        self.table.tag_configure("sel",   background=ACCENTL)

        t_scroll_y = ttk.Scrollbar(table_frame, orient="vertical",
                                    command=self.table.yview)
        t_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal",
                                    command=self.table.xview)
        self.table.configure(yscrollcommand=t_scroll_y.set,
                             xscrollcommand=t_scroll_x.set)
        t_scroll_y.pack(side="right", fill="y")
        t_scroll_x.pack(side="bottom", fill="x")
        self.table.pack(fill="both", expand=True)
        self.table.bind("<<TreeviewSelect>>", self._on_row_select)
        self.table.bind("<Return>", lambda e: self._focus_edit())

        # Edit panel (below table)
        tk.Frame(right_side, bg=BORDER, height=1).pack(fill="x")
        self._edit_panel = tk.Frame(right_side, bg=BG, padx=16, pady=10)
        self._edit_panel.pack(fill="x")
        self._build_edit_panel()

        # Status bar
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        sbar = tk.Frame(self, bg=BG, padx=16, pady=5)
        sbar.pack(fill="x")
        tk.Label(sbar, textvariable=self._status, bg=BG,
                 fg=TEXT2, font=FONTS, anchor="w").pack(side="left")
        self._ai_status = tk.Label(sbar, text="", bg=BG, fg=ACCENT, font=FONTS)
        self._ai_status.pack(side="right")

    def _build_edit_panel(self):
        p = self._edit_panel

        # Key label
        key_row = tk.Frame(p, bg=BG)
        key_row.pack(fill="x", pady=(0, 8))
        self._key_lbl = tk.Label(key_row, text="Select a string to edit",
                                  bg=BG, fg=TEXT2, font=FONTM)
        self._key_lbl.pack(side="left")

        # Source + translation side by side
        cols = tk.Frame(p, bg=BG)
        cols.pack(fill="x")
        cols.columnconfigure(0, weight=1)
        cols.columnconfigure(1, weight=1)

        # Source
        src_frame = tk.Frame(cols, bg=BG)
        src_frame.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        tk.Label(src_frame, text="Source", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        self._src_text = tk.Text(src_frame, height=3, font=FONT,
                                  bg=BG2, fg=TEXT, relief="flat",
                                  wrap="word", state="disabled",
                                  padx=8, pady=6, cursor="arrow")
        self._src_text.pack(fill="x")

        # Translation
        tgt_frame = tk.Frame(cols, bg=BG)
        tgt_frame.grid(row=0, column=1, sticky="ew")
        tgt_frame.columnconfigure(0, weight=1)

        tgt_top = tk.Frame(tgt_frame, bg=BG)
        tgt_top.pack(fill="x")
        tk.Label(tgt_top, text="Translation", bg=BG, fg=TEXT2, font=FONTS).pack(side="left")
        self._ai_btn = ttk.Button(tgt_top, text="AI suggest",
                                   command=self._ai_single,
                                   style="Icon.TButton")
        self._ai_btn.pack(side="right")
        self._save_btn = ttk.Button(tgt_top, text="Save  ↵",
                                     command=self._save_edit,
                                     style="Icon.TButton")
        self._save_btn.pack(side="right", padx=(0, 4))

        self._tgt_text = tk.Text(tgt_frame, height=3, font=FONT,
                                  bg="#FFFFFF", fg=TEXT, relief="flat",
                                  wrap="word", padx=8, pady=6,
                                  insertbackground=ACCENT,
                                  highlightthickness=1,
                                  highlightbackground=BORDER,
                                  highlightcolor=ACCENT)
        self._tgt_text.pack(fill="x")
        self._tgt_text.bind("<Control-Return>", lambda e: self._save_edit())
        self._tgt_text.bind("<Escape>", lambda e: self.table.focus_set())

        self._cur_key = None  # key being edited

    # ── Data helpers ──────────────────────────────────────────────────────

    def _load_records(self, records, translations, path_label):
        self.records      = records
        self.translations = translations
        self.changed      = False
        self._cur_key     = None
        self._cur_grp     = None
        self._build_groups()
        self._refresh_table()
        self._update_progress()
        self._file_label.config(text=path_label)
        self._status.set(f"Loaded {len(records):,} strings  ·  "
                         f"{len(translations):,} translated")

    def _build_groups(self):
        self.grp_tv.delete(*self.grp_tv.get_children())
        self._grp_ids = {}
        groups: dict[str, int] = {}
        for r in self.records:
            parts = r.key.split('/')
            g = parts[0] if len(parts) == 1 else '/'.join(parts[:2])
            groups[g] = groups.get(g, 0) + 1
        all_id = self.grp_tv.insert("", "end",
                                     text=f"All  ({len(self.records):,})")
        self._grp_ids[all_id] = None
        for g in sorted(groups):
            iid = self.grp_tv.insert("", "end", text=f"{g}  ({groups[g]})")
            self._grp_ids[iid] = g
        self.grp_tv.selection_set(all_id)

    def _on_group_select(self, _=None):
        sel = self.grp_tv.selection()
        if sel:
            self._cur_grp = self._grp_ids.get(sel[0])
            self._refresh_table()

    def _refresh_table(self):
        q    = self._search_q.get().lower()
        emp  = self._show_empty.get()
        grp  = self._cur_grp

        # Gather matching rows
        rows = []
        for r in self.records:
            if grp is not None:
                rg = r.key.split('/')[0] if '/' not in r.key else '/'.join(r.key.split('/')[:2])
                if rg != grp:
                    continue
            tval = self.translations.get(r.key, "")
            if emp and tval.strip():
                continue
            if q and q not in r.key.lower() and q not in r.value.lower():
                continue
            rows.append((r, tval))

        # Rebuild treeview — treeview handles 13k rows without hanging
        self.table.delete(*self.table.get_children())
        self._tv_items = []
        for r, tval in rows:
            done = bool(tval.strip())
            tag  = "done" if done else "todo"
            dot  = "●" if done else "○"
            iid  = self.table.insert("", "end",
                                      iid=r.key,
                                      values=(r.value, tval, dot),
                                      tags=(tag,))
            self._tv_items.append(iid)

    def _on_row_select(self, _=None):
        sel = self.table.selection()
        if not sel:
            return
        key = sel[0]  # iid == key
        rec = next((r for r in self.records if r.key == key), None)
        if not rec:
            return
        self._cur_key = key
        self._key_lbl.config(text=key)

        self._src_text.config(state="normal")
        self._src_text.delete("1.0", "end")
        self._src_text.insert("end", rec.value)
        self._src_text.config(state="disabled")

        self._tgt_text.delete("1.0", "end")
        self._tgt_text.insert("end", self.translations.get(key, ""))

    def _focus_edit(self):
        self._tgt_text.focus_set()
        self._tgt_text.mark_set("insert", "end")

    def _save_edit(self, _=None):
        if not self._cur_key:
            return
        val = self._tgt_text.get("1.0", "end-1c").strip()
        key = self._cur_key
        if val:
            self.translations[key] = val
        elif key in self.translations:
            del self.translations[key]
        self.changed = True
        # Update the table row in-place (fast — no full rebuild)
        dot = "●" if val else "○"
        tag = "done" if val else "todo"
        try:
            self.table.item(key, values=(
                self.table.item(key)["values"][0], val, dot), tags=(tag,))
        except tk.TclError:
            pass
        self._update_progress()
        # Move to next untranslated row
        self._advance_to_next()

    def _advance_to_next(self):
        """Select next row in the table after saving."""
        items = self.table.get_children()
        if not items:
            return
        cur  = self._cur_key
        if cur in items:
            idx  = list(items).index(cur)
            next_idx = (idx + 1) % len(items)
            nxt = items[next_idx]
            self.table.selection_set(nxt)
            self.table.see(nxt)
            self._on_row_select()

    def _update_progress(self):
        total = len(self.records)
        done  = len(self.translations)
        if not total:
            self._progress.set(""); self._prog_bar["value"] = 0; return
        pct = int(done / total * 100)
        self._progress.set(f"{done:,} / {total:,}  ({pct}%)")
        self._prog_bar["value"] = pct

    # ── File actions ──────────────────────────────────────────────────────

    def _import_lang(self):
        p = filedialog.askopenfilename(
            title="Import .lang file (source language)",
            filetypes=[("Lang files", "*.lang *.lngp"), ("All files", "*.*")])
        if not p:
            return
        self._status.set(f"Importing {Path(p).name}…")
        self.update_idletasks()
        try:
            records = load_lang(Path(p))
        except Exception as e:
            messagebox.showerror("Import error", str(e)); return

        # Suggest where to save the CSV
        default = Path(p).with_suffix('.csv')
        csv_out = filedialog.asksaveasfilename(
            title="Save working CSV",
            initialfile=default.name,
            initialdir=str(default.parent),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not csv_out:
            return

        save_csv(Path(csv_out), records, {})
        self.csv_path = Path(csv_out)
        self._load_records(records, {}, Path(csv_out).name)
        self._status.set(
            f"Imported {len(records):,} strings from {Path(p).name}  ·  "
            f"Working CSV saved as {Path(csv_out).name}")

    def _open_csv(self):
        p = filedialog.askopenfilename(
            title="Open working CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not p:
            return
        try:
            records, translations = load_csv(Path(p))
        except Exception as e:
            messagebox.showerror("Open error", str(e)); return
        self.csv_path = Path(p)
        self._load_records(records, translations, Path(p).name)

    def _save_csv_action(self):
        if not self.records:
            messagebox.showinfo("Nothing to save", "Import a .lang file first."); return
        if self.csv_path:
            self._do_save_csv(self.csv_path)
        else:
            p = filedialog.asksaveasfilename(
                title="Save CSV",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
            if p:
                self.csv_path = Path(p)
                self._do_save_csv(self.csv_path)

    def _do_save_csv(self, path: Path):
        # Flush current edit
        if self._cur_key:
            val = self._tgt_text.get("1.0", "end-1c").strip()
            if val: self.translations[self._cur_key] = val
        try:
            save_csv(path, self.records, self.translations)
            self.changed = False
            self._status.set(f"Saved  {path.name}  ·  "
                             f"{len(self.translations):,} translations written")
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def _pack_lang(self):
        if not self.records:
            messagebox.showinfo("Nothing to pack", "Open a CSV or import a .lang file first.")
            return
        if self.changed or self._cur_key:
            self._do_save_csv(self.csv_path) if self.csv_path else None

        lang = self.tgt_lang.get().strip()
        if not lang:
            messagebox.showwarning("No language", "Enter a target language name.")
            return

        default_name = f"{lang.lower()}.lang"
        init_dir = str(self.csv_path.parent) if self.csv_path else "."
        out = filedialog.asksaveasfilename(
            title="Pack to .lang binary",
            initialfile=default_name,
            initialdir=init_dir,
            defaultextension=".lang",
            filetypes=[("Lang files", "*.lang"), ("All files", "*.*")])
        if not out:
            return
        try:
            data = build_lang(self.records, self.translations)
            Path(out).write_bytes(data)
            done = len(self.translations)
            total = len(self.records)
            self._status.set(
                f"Packed  {Path(out).name}  ·  {done:,}/{total:,} strings  ·  "
                f"{len(data):,} bytes")
            messagebox.showinfo("Packed",
                f"Saved {Path(out).name}\n\n"
                f"Next steps:\n"
                f"1. Copy {Path(out).name} to the game's  Localizations\\  folder\n"
                f"2. Open Storyteller → select your mod → click Pack\n"
                f"3. Launch the game and select your language")
        except Exception as e:
            messagebox.showerror("Pack error", str(e))

    # ── AI ────────────────────────────────────────────────────────────────

    def _ai_single(self):
        if self.ai_busy:
            return
        key = self.api_key.get().strip()
        if not key:
            self._open_settings(); return
        if not self._cur_key:
            return
        rec = next((r for r in self.records if r.key == self._cur_key), None)
        if not rec or not rec.value.strip():
            return
        self.ai_busy = True
        self._ai_status.config(text="AI translating…")
        src  = (self.csv_path.stem if self.csv_path else "English")
        tgt  = self.tgt_lang.get()
        text = rec.value
        def work():
            try:
                res = ai_translate(text, src, tgt, key)
                self.after(0, lambda: self._apply_ai(res))
            except Exception as e:
                self.after(0, lambda: self._status.set(f"AI error: {e}"))
            finally:
                self.after(0, lambda: (self._ai_status.config(text=""),
                                       setattr(self, "ai_busy", False)))
        threading.Thread(target=work, daemon=True).start()

    def _apply_ai(self, result):
        self._tgt_text.delete("1.0", "end")
        self._tgt_text.insert("end", result)
        self._tgt_text.focus_set()

    def _ai_batch(self):
        key = self.api_key.get().strip()
        if not key:
            self._open_settings(); return
        untrans = [r for r in self.records
                   if r.value.strip() and r.key not in self.translations]
        if not untrans:
            messagebox.showinfo("Done", "All strings already translated!"); return
        if not messagebox.askyesno("Batch AI",
            f"Auto-translate {len(untrans):,} empty strings?\n"
            "This will use API credits. It may take several minutes."): return
        self.ai_busy = True
        src = self.csv_path.stem if self.csv_path else "English"
        tgt = self.tgt_lang.get()
        def work():
            done = 0
            for rec in untrans:
                try:
                    res = ai_translate(rec.value, src, tgt, key)
                    self.translations[rec.key] = res; done += 1
                    self.after(0, lambda d=done:
                        self._ai_status.config(text=f"AI {d:,}/{len(untrans):,}"))
                except Exception as e:
                    self.after(0, lambda ex=e: self._status.set(f"AI error: {ex}")); break
            self.after(0, self._ai_done, done, len(untrans))
        threading.Thread(target=work, daemon=True).start()

    def _ai_done(self, done, total):
        self.ai_busy = False; self.changed = True
        self._ai_status.config(text="")
        self._refresh_table(); self._update_progress()
        self._status.set(f"AI batch: {done:,}/{total:,} translated.")

    # ── Settings ──────────────────────────────────────────────────────────

    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings"); win.geometry("460x220")
        win.configure(bg=BG); win.resizable(False, False); win.grab_set()

        tk.Label(win, text="Settings", bg=BG, fg=TEXT,
                 font=("Segoe UI", 13, "bold")).pack(pady=(20, 16), padx=24, anchor="w")

        f = tk.Frame(win, bg=BG); f.pack(padx=24, fill="x")
        tk.Label(f, text="Anthropic API key", bg=BG, fg=TEXT2, font=FONTS).pack(anchor="w")
        tk.Label(f, text="For AI translation — get a free key at console.anthropic.com",
                 bg=BG, fg=TEXT2, font=("Segoe UI", 8)).pack(anchor="w", pady=(0,4))
        ae = ttk.Entry(f, textvariable=self.api_key, width=52, show="*"); ae.pack(fill="x")
        sv = tk.BooleanVar()
        tk.Checkbutton(f, text="Show key", variable=sv,
                       command=lambda: ae.config(show="" if sv.get() else "*"),
                       bg=BG, fg=TEXT2, font=FONTS, activebackground=BG).pack(anchor="w", pady=4)

        bf = tk.Frame(win, bg=BG); bf.pack(pady=12, padx=24, anchor="w")
        def on_save():
            self.cfg.update({"api_key": self.api_key.get(), "tgt_lang": self.tgt_lang.get()})
            save_cfg(self.cfg); win.destroy()
        ttk.Button(bf, text="Save", command=on_save,
                   style="Primary.TButton").pack(side="left", padx=(0,8))
        ttk.Button(bf, text="Batch translate all empty",
                   command=lambda: [on_save(), self._ai_batch()],
                   style="Ghost.TButton").pack(side="left")

    def _on_close(self):
        if self.changed:
            r = messagebox.askyesnocancel("Unsaved changes",
                "Save CSV before quitting?")
            if r is None: return
            if r and self.csv_path: self._do_save_csv(self.csv_path)
        self.cfg.update({"tgt_lang": self.tgt_lang.get(),
                         "api_key":  self.api_key.get()})
        save_cfg(self.cfg); self.destroy()

    # ── Menubar (keyboard shortcuts) ──────────────────────────────────────
    def _build_menu(self):
        m = tk.Menu(self, tearoff=0, bg=BG, fg=TEXT, relief="flat")
        self.config(menu=m)
        file_m = tk.Menu(m, tearoff=0, bg=BG, fg=TEXT)
        m.add_cascade(label="File", menu=file_m)
        file_m.add_command(label="Import .lang…", command=self._import_lang, accelerator="Ctrl+I")
        file_m.add_command(label="Open CSV…",     command=self._open_csv,    accelerator="Ctrl+O")
        file_m.add_command(label="Save CSV",       command=self._save_csv_action, accelerator="Ctrl+S")
        file_m.add_separator()
        file_m.add_command(label="Pack to .lang…", command=self._pack_lang,  accelerator="Ctrl+P")
        file_m.add_separator()
        file_m.add_command(label="Settings…",      command=self._open_settings)
        self.bind_all("<Control-i>", lambda e: self._import_lang())
        self.bind_all("<Control-o>", lambda e: self._open_csv())
        self.bind_all("<Control-s>", lambda e: self._save_csv_action())
        self.bind_all("<Control-p>", lambda e: self._pack_lang())
        self.bind_all("<Control-Return>", lambda e: self._save_edit())


if __name__ == "__main__":
    app = App()
    app._build_menu()
    app.mainloop()
