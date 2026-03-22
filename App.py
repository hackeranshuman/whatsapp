"""
WhatsApp Bulk Messenger - Desktop Application
A professional Tkinter-based GUI for sending WhatsApp messages via pywhatkit or Twilio.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time
import pandas as pd
import re
import os
import sys
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# THEME & PALETTE
# ─────────────────────────────────────────────────────────────────────────────
PALETTE = {
    "bg":           "#0F1117",
    "surface":      "#1A1D2E",
    "surface2":     "#242740",
    "accent":       "#25D366",   # WhatsApp green
    "accent2":      "#128C7E",
    "accent_dim":   "#1a9e52",
    "danger":       "#FF5C5C",
    "warning":      "#FFB347",
    "text":         "#E8EAF0",
    "text_dim":     "#8B8FA8",
    "border":       "#2E3150",
    "entry_bg":     "#151825",
    "success":      "#25D366",
}

FONT_TITLE   = ("Segoe UI", 22, "bold")
FONT_SECTION = ("Segoe UI", 11, "bold")
FONT_LABEL   = ("Segoe UI", 10)
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)
FONT_BTN     = ("Segoe UI", 10, "bold")


# ─────────────────────────────────────────────────────────────────────────────
# HELPER – phone validation
# ─────────────────────────────────────────────────────────────────────────────
def normalize_phone(phone: str) -> str:
    """Return digits-only phone number; raise ValueError if invalid."""
    digits = re.sub(r"[^\d+]", "", str(phone).strip())
    digits = digits.lstrip("+")
    if len(digits) < 7:
        raise ValueError(f"Phone number too short: {phone!r}")
    return digits


def format_message(template: str, name: str) -> str:
    return template.replace("{name}", name)


# ─────────────────────────────────────────────────────────────────────────────
# SENDER BACKENDS
# ─────────────────────────────────────────────────────────────────────────────
class PyWhatKitSender:
    """Sends via WhatsApp Web using pywhatkit (opens browser)."""

    def send(self, phone: str, message: str, wait: int = 20) -> None:
        import pywhatkit as kit
        phone_e164 = "+" + normalize_phone(phone)
        now = datetime.now()
        # Schedule 1 minute ahead so pywhatkit has time to open the browser
        send_time_min = now.minute + 1
        send_time_hour = now.hour
        if send_time_min >= 60:
            send_time_min -= 60
            send_time_hour = (send_time_hour + 1) % 24

        kit.sendwhatmsg(
            phone_e164,
            message,
            send_time_hour,
            send_time_min,
            wait_time=wait,
            tab_close=True,
            close_time=3,
        )


class TwilioSender:
    """Sends via Twilio WhatsApp API (production-grade)."""

    def __init__(self, account_sid: str, auth_token: str, from_number: str):
        from twilio.rest import Client
        self.client = Client(account_sid, auth_token)
        # Ensure whatsapp: prefix
        self.from_number = (
            from_number if from_number.startswith("whatsapp:")
            else f"whatsapp:{from_number}"
        )

    def send(self, phone: str, message: str, **_) -> None:
        to = "whatsapp:+" + normalize_phone(phone)
        self.client.messages.create(
            body=message,
            from_=self.from_number,
            to=to,
        )


# ─────────────────────────────────────────────────────────────────────────────
# STYLED WIDGETS
# ─────────────────────────────────────────────────────────────────────────────
class StyledFrame(tk.Frame):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["surface"])
        kw.setdefault("relief", "flat")
        super().__init__(parent, **kw)


class SectionLabel(tk.Label):
    def __init__(self, parent, text, **kw):
        kw.setdefault("font", FONT_SECTION)
        kw.setdefault("fg", PALETTE["accent"])
        kw.setdefault("bg", PALETTE["surface"])
        super().__init__(parent, text=text, **kw)


class StyledEntry(tk.Entry):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["entry_bg"])
        kw.setdefault("fg", PALETTE["text"])
        kw.setdefault("insertbackground", PALETTE["accent"])
        kw.setdefault("relief", "flat")
        kw.setdefault("font", FONT_LABEL)
        kw.setdefault("highlightthickness", 1)
        kw.setdefault("highlightcolor", PALETTE["accent"])
        kw.setdefault("highlightbackground", PALETTE["border"])
        super().__init__(parent, **kw)


class AccentButton(tk.Button):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["accent"])
        kw.setdefault("fg", "#0F1117")
        kw.setdefault("activebackground", PALETTE["accent2"])
        kw.setdefault("activeforeground", PALETTE["text"])
        kw.setdefault("relief", "flat")
        kw.setdefault("font", FONT_BTN)
        kw.setdefault("cursor", "hand2")
        kw.setdefault("padx", 14)
        kw.setdefault("pady", 7)
        super().__init__(parent, **kw)


class DangerButton(tk.Button):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["danger"])
        kw.setdefault("fg", "#fff")
        kw.setdefault("activebackground", "#cc4444")
        kw.setdefault("activeforeground", "#fff")
        kw.setdefault("relief", "flat")
        kw.setdefault("font", FONT_BTN)
        kw.setdefault("cursor", "hand2")
        kw.setdefault("padx", 14)
        kw.setdefault("pady", 7)
        super().__init__(parent, **kw)


class GhostButton(tk.Button):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["surface2"])
        kw.setdefault("fg", PALETTE["text"])
        kw.setdefault("activebackground", PALETTE["border"])
        kw.setdefault("activeforeground", PALETTE["text"])
        kw.setdefault("relief", "flat")
        kw.setdefault("font", FONT_BTN)
        kw.setdefault("cursor", "hand2")
        kw.setdefault("padx", 14)
        kw.setdefault("pady", 7)
        super().__init__(parent, **kw)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class WhatsAppMessengerApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("WhatsApp Bulk Messenger")
        self.geometry("1100x800")
        self.minsize(900, 650)
        self.configure(bg=PALETTE["bg"])

        # State
        self.contacts: list[dict] = []          # {"name": ..., "phone": ...}
        self.sender_mode = tk.StringVar(value="pywhatkit")
        self.is_sending   = False
        self.stop_flag    = threading.Event()

        # Twilio credentials
        self.twilio_sid   = tk.StringVar()
        self.twilio_token = tk.StringVar()
        self.twilio_from  = tk.StringVar()

        # Build UI
        self._build_header()
        self._build_main()
        self._build_status_bar()

        # Apply ttk theme overrides
        self._apply_ttk_styles()

    # ── ttk style ─────────────────────────────────────────────────────────────
    def _apply_ttk_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background=PALETTE["entry_bg"],
                        foreground=PALETTE["text"],
                        rowheight=26,
                        fieldbackground=PALETTE["entry_bg"],
                        bordercolor=PALETTE["border"],
                        font=FONT_LABEL)
        style.configure("Treeview.Heading",
                        background=PALETTE["surface2"],
                        foreground=PALETTE["accent"],
                        font=FONT_SECTION,
                        relief="flat")
        style.map("Treeview",
                  background=[("selected", PALETTE["accent2"])],
                  foreground=[("selected", "#fff")])
        style.configure("TNotebook",
                        background=PALETTE["bg"],
                        borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=PALETTE["surface"],
                        foreground=PALETTE["text_dim"],
                        padding=(16, 8),
                        font=FONT_SECTION)
        style.map("TNotebook.Tab",
                  background=[("selected", PALETTE["surface2"])],
                  foreground=[("selected", PALETTE["accent"])])
        style.configure("TProgressbar",
                        troughcolor=PALETTE["surface2"],
                        background=PALETTE["accent"],
                        borderwidth=0,
                        thickness=6)
        style.configure("Vertical.TScrollbar",
                        background=PALETTE["surface2"],
                        troughcolor=PALETTE["surface"],
                        borderwidth=0,
                        arrowcolor=PALETTE["text_dim"])
        style.configure("Horizontal.TScrollbar",
                        background=PALETTE["surface2"],
                        troughcolor=PALETTE["surface"],
                        borderwidth=0,
                        arrowcolor=PALETTE["text_dim"])

    # ── Header ─────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=PALETTE["surface"], pady=0)
        hdr.pack(fill="x", side="top")

        # Green left accent bar
        tk.Frame(hdr, bg=PALETTE["accent"], width=5).pack(side="left", fill="y")

        inner = tk.Frame(hdr, bg=PALETTE["surface"], padx=20, pady=14)
        inner.pack(side="left", fill="both", expand=True)

        tk.Label(inner, text="📱  WhatsApp Bulk Messenger",
                 font=FONT_TITLE, fg=PALETTE["text"],
                 bg=PALETTE["surface"]).pack(side="left")

        tk.Label(inner,
                 text="Send personalised messages to all your clients",
                 font=FONT_SMALL, fg=PALETTE["text_dim"],
                 bg=PALETTE["surface"]).pack(side="left", padx=(18, 0), pady=(6, 0))

        # Stats badges
        self.badge_total = self._badge(inner, "0 Contacts", PALETTE["accent2"])
        self.badge_total.pack(side="right", padx=(0, 6))

    def _badge(self, parent, text, color):
        f = tk.Frame(parent, bg=color, padx=10, pady=4)
        lbl = tk.Label(f, text=text, font=FONT_SMALL, fg="#fff", bg=color)
        lbl.pack()
        f._label = lbl
        return f

    # ── Main layout ────────────────────────────────────────────────────────────
    def _build_main(self):
        main = tk.Frame(self, bg=PALETTE["bg"])
        main.pack(fill="both", expand=True, padx=14, pady=(10, 0))

        # Left panel (inputs)
        left = StyledFrame(main, width=420)
        left.pack(side="left", fill="y", padx=(0, 8), pady=0)
        left.pack_propagate(False)
        self._build_left_panel(left)

        # Right panel (contacts + log)
        right = tk.Frame(main, bg=PALETTE["bg"])
        right.pack(side="left", fill="both", expand=True)
        self._build_right_panel(right)

    # ── Left panel ─────────────────────────────────────────────────────────────
    def _build_left_panel(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=2, pady=2)

        contacts_tab = StyledFrame(nb)
        message_tab  = StyledFrame(nb)
        settings_tab = StyledFrame(nb)

        nb.add(contacts_tab, text="  👥 Contacts  ")
        nb.add(message_tab,  text="  ✉️ Message   ")
        nb.add(settings_tab, text="  ⚙️ Settings  ")

        self._build_contacts_tab(contacts_tab)
        self._build_message_tab(message_tab)
        self._build_settings_tab(settings_tab)

    # ── Contacts tab ───────────────────────────────────────────────────────────
    def _build_contacts_tab(self, parent):
        pad = dict(padx=16, pady=6)

        # CSV section
        SectionLabel(parent, "📂  Import from CSV").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)

        csv_info = tk.Label(parent,
                            text="CSV must have columns: name, phone_number",
                            font=FONT_SMALL, fg=PALETTE["text_dim"],
                            bg=PALETTE["surface"], wraplength=360, justify="left")
        csv_info.pack(anchor="w", padx=16, pady=(4, 8))

        btn_row = tk.Frame(parent, bg=PALETTE["surface"])
        btn_row.pack(fill="x", padx=16, pady=(0, 10))
        AccentButton(btn_row, text="📁  Browse CSV",
                     command=self._load_csv).pack(side="left")
        GhostButton(btn_row, text="Download Template",
                    command=self._download_template).pack(side="left", padx=(8, 0))

        # Manual entry section
        SectionLabel(parent, "✏️  Manual Entry").pack(anchor="w", padx=16, pady=(10, 2))
        self._divider(parent)

        form = tk.Frame(parent, bg=PALETTE["surface"])
        form.pack(fill="x", padx=16, pady=8)

        tk.Label(form, text="Name", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]).grid(row=0, column=0, sticky="w", pady=3)
        self.ent_name = StyledEntry(form, width=28)
        self.ent_name.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        tk.Label(form, text="Phone (with country code, e.g. +91xxxxxxxxxx)",
                 font=FONT_SMALL, fg=PALETTE["text_dim"],
                 bg=PALETTE["surface"]).grid(row=2, column=0, sticky="w", pady=3)
        self.ent_phone = StyledEntry(form, width=28)
        self.ent_phone.grid(row=3, column=0, sticky="ew")
        form.columnconfigure(0, weight=1)

        add_row = tk.Frame(parent, bg=PALETTE["surface"])
        add_row.pack(fill="x", padx=16, pady=(4, 0))
        AccentButton(add_row, text="＋  Add Contact",
                     command=self._add_manual_contact).pack(side="left")
        DangerButton(add_row, text="✕  Clear All",
                     command=self._clear_contacts).pack(side="left", padx=(8, 0))

    # ── Message tab ────────────────────────────────────────────────────────────
    def _build_message_tab(self, parent):
        pad = dict(padx=16, pady=6)

        SectionLabel(parent, "💬  Message Template").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)

        ph_info = tk.Label(parent,
                           text="Use {name} as a placeholder for the recipient's name.",
                           font=FONT_SMALL, fg=PALETTE["text_dim"],
                           bg=PALETTE["surface"], wraplength=360, justify="left")
        ph_info.pack(anchor="w", padx=16, pady=(4, 6))

        # Template quick-inserts
        qi_row = tk.Frame(parent, bg=PALETTE["surface"])
        qi_row.pack(fill="x", padx=16, pady=(0, 6))
        tk.Label(qi_row, text="Insert:", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]).pack(side="left")
        for ph in ["{name}"]:
            GhostButton(qi_row, text=ph,
                        command=lambda p=ph: self._insert_placeholder(p),
                        padx=8, pady=3).pack(side="left", padx=(6, 0))

        self.txt_message = scrolledtext.ScrolledText(
            parent, height=10,
            bg=PALETTE["entry_bg"], fg=PALETTE["text"],
            insertbackground=PALETTE["accent"],
            relief="flat", font=FONT_LABEL,
            wrap="word",
            highlightthickness=1,
            highlightcolor=PALETTE["accent"],
            highlightbackground=PALETTE["border"],
        )
        self.txt_message.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # Preview
        SectionLabel(parent, "👁  Preview").pack(anchor="w", padx=16, pady=(6, 2))
        self._divider(parent)

        preview_row = tk.Frame(parent, bg=PALETTE["surface"])
        preview_row.pack(fill="x", padx=16, pady=6)
        GhostButton(preview_row, text="Preview for first contact",
                    command=self._preview_message).pack(side="left")

        self.lbl_preview = tk.Label(parent,
                                    text="(preview will appear here)",
                                    font=FONT_SMALL, fg=PALETTE["text_dim"],
                                    bg=PALETTE["surface2"],
                                    wraplength=360, justify="left",
                                    padx=12, pady=10, anchor="w")
        self.lbl_preview.pack(fill="x", padx=16, pady=(0, 10))

        # Delay
        SectionLabel(parent, "⏱  Delay Between Messages").pack(anchor="w", padx=16, pady=(6, 2))
        self._divider(parent)

        delay_row = tk.Frame(parent, bg=PALETTE["surface"])
        delay_row.pack(fill="x", padx=16, pady=8)
        tk.Label(delay_row, text="Delay (seconds):", font=FONT_LABEL,
                 fg=PALETTE["text"], bg=PALETTE["surface"]).pack(side="left")
        self.spin_delay = tk.Spinbox(delay_row, from_=5, to=120, width=5,
                                     bg=PALETTE["entry_bg"], fg=PALETTE["text"],
                                     buttonbackground=PALETTE["surface2"],
                                     relief="flat", font=FONT_LABEL,
                                     highlightthickness=1,
                                     highlightbackground=PALETTE["border"])
        self.spin_delay.pack(side="left", padx=(8, 0))
        self.spin_delay.delete(0, "end")
        self.spin_delay.insert(0, "15")

    # ── Settings tab ───────────────────────────────────────────────────────────
    def _build_settings_tab(self, parent):
        pad = dict(padx=16, pady=6)

        SectionLabel(parent, "🔌  Sending Method").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)

        modes = [
            ("pywhatkit", "📲  PyWhatKit  (WhatsApp Web — No API key needed)"),
            ("twilio",    "☁️   Twilio API  (Production — Requires credentials)"),
        ]
        for val, label in modes:
            rb = tk.Radiobutton(parent, text=label,
                                variable=self.sender_mode, value=val,
                                bg=PALETTE["surface"], fg=PALETTE["text"],
                                selectcolor=PALETTE["surface2"],
                                activebackground=PALETTE["surface"],
                                activeforeground=PALETTE["accent"],
                                font=FONT_LABEL, anchor="w",
                                command=self._on_mode_change)
            rb.pack(fill="x", padx=20, pady=4)

        # PyWhatKit note
        self.frame_pywhatkit_note = tk.Frame(parent, bg=PALETTE["surface2"],
                                              padx=14, pady=10)
        self.frame_pywhatkit_note.pack(fill="x", padx=16, pady=(8, 0))
        tk.Label(self.frame_pywhatkit_note,
                 text="ℹ️  Requires WhatsApp Web to be open and logged in.\n"
                      "Each message will briefly open a browser tab.",
                 font=FONT_SMALL, fg=PALETTE["warning"],
                 bg=PALETTE["surface2"], justify="left").pack(anchor="w")

        # Twilio creds
        self.frame_twilio = tk.Frame(parent, bg=PALETTE["surface"])
        self.frame_twilio.pack(fill="x", padx=16, pady=(10, 0))

        SectionLabel(self.frame_twilio, "☁️  Twilio Credentials").pack(anchor="w", pady=(6, 2))
        self._divider(self.frame_twilio)

        fields = [
            ("Account SID",   self.twilio_sid,   False),
            ("Auth Token",    self.twilio_token, True),
            ("From Number\n(e.g. +14155238886)", self.twilio_from, False),
        ]
        for label, var, secret in fields:
            tk.Label(self.frame_twilio, text=label, font=FONT_SMALL,
                     fg=PALETTE["text_dim"], bg=PALETTE["surface"]).pack(anchor="w", pady=(6, 0))
            ent = StyledEntry(self.frame_twilio, textvariable=var,
                              show="●" if secret else "")
            ent.pack(fill="x", pady=(2, 0))

        AccentButton(self.frame_twilio, text="✅  Test Twilio Connection",
                     command=self._test_twilio).pack(pady=10)

        self.frame_twilio.pack_forget()   # hidden until twilio mode selected

    # ── Right panel (contacts table + log) ────────────────────────────────────
    def _build_right_panel(self, parent):
        # Contacts table
        top = StyledFrame(parent)
        top.pack(fill="both", expand=True, pady=(0, 8))

        hdr = tk.Frame(top, bg=PALETTE["surface"])
        hdr.pack(fill="x", padx=16, pady=(14, 4))
        SectionLabel(hdr, "👥  Contact List").pack(side="left")
        self.lbl_count = tk.Label(hdr, text="0 contacts",
                                  font=FONT_SMALL, fg=PALETTE["text_dim"],
                                  bg=PALETTE["surface"])
        self.lbl_count.pack(side="left", padx=10)

        self._divider(top)

        tree_frame = tk.Frame(top, bg=PALETTE["surface"])
        tree_frame.pack(fill="both", expand=True, padx=16, pady=8)

        cols = ("no", "name", "phone", "status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
        for col, w, heading in [
            ("no",     50,  "#"),
            ("name",  200,  "Name"),
            ("phone", 200,  "Phone"),
            ("status",130,  "Status"),
        ]:
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=w, minwidth=40, anchor="w" if col != "no" else "center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Context menu on tree
        self.ctx_menu = tk.Menu(self, tearoff=0, bg=PALETTE["surface2"],
                                fg=PALETTE["text"], activebackground=PALETTE["accent2"],
                                font=FONT_SMALL)
        self.ctx_menu.add_command(label="Remove selected", command=self._remove_selected)
        self.tree.bind("<Button-3>", self._show_ctx_menu)

        # Send controls
        send_bar = tk.Frame(top, bg=PALETTE["surface"])
        send_bar.pack(fill="x", padx=16, pady=(0, 14))

        AccentButton(send_bar, text="▶  Send to All",
                     command=lambda: self._start_sending("all")).pack(side="left")
        GhostButton(send_bar, text="▶  Send to Selected",
                    command=lambda: self._start_sending("selected")).pack(side="left", padx=(8, 0))
        DangerButton(send_bar, text="■  Stop",
                     command=self._stop_sending).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(send_bar, orient="horizontal",
                                        mode="determinate", length=180)
        self.progress.pack(side="right", padx=(0, 0))
        self.lbl_progress = tk.Label(send_bar, text="",
                                     font=FONT_SMALL, fg=PALETTE["text_dim"],
                                     bg=PALETTE["surface"])
        self.lbl_progress.pack(side="right", padx=(0, 8))

        # Log section
        log_frame = StyledFrame(parent)
        log_frame.pack(fill="both", expand=False, pady=(0, 0))

        log_hdr = tk.Frame(log_frame, bg=PALETTE["surface"])
        log_hdr.pack(fill="x", padx=16, pady=(10, 4))
        SectionLabel(log_hdr, "📋  Activity Log").pack(side="left")
        GhostButton(log_hdr, text="Clear Log", padx=8, pady=3,
                    command=self._clear_log).pack(side="right")

        self._divider(log_frame)

        self.txt_log = scrolledtext.ScrolledText(
            log_frame, height=8,
            bg=PALETTE["entry_bg"], fg=PALETTE["text"],
            font=FONT_MONO, state="disabled",
            relief="flat",
            highlightthickness=0,
        )
        self.txt_log.pack(fill="both", expand=True, padx=16, pady=(6, 12))
        # Tag colours
        self.txt_log.tag_config("success", foreground=PALETTE["success"])
        self.txt_log.tag_config("error",   foreground=PALETTE["danger"])
        self.txt_log.tag_config("info",    foreground=PALETTE["text_dim"])
        self.txt_log.tag_config("warn",    foreground=PALETTE["warning"])

    # ── Status bar ─────────────────────────────────────────────────────────────
    def _build_status_bar(self):
        bar = tk.Frame(self, bg=PALETTE["surface2"], height=28)
        bar.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(bar, text="Ready  •  Select a sending method and load contacts",
                                   font=FONT_SMALL, fg=PALETTE["text_dim"],
                                   bg=PALETTE["surface2"])
        self.lbl_status.pack(side="left", padx=14)

        tk.Label(bar, text="WhatsApp Bulk Messenger v1.0",
                 font=FONT_SMALL, fg=PALETTE["border"],
                 bg=PALETTE["surface2"]).pack(side="right", padx=14)

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _divider(self, parent):
        tk.Frame(parent, bg=PALETTE["border"], height=1).pack(fill="x", padx=16)

    def _set_status(self, msg: str):
        self.lbl_status.config(text=msg)

    def _log(self, msg: str, tag: str = "info"):
        self.txt_log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {msg}\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _clear_log(self):
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")

    def _insert_placeholder(self, ph: str):
        self.txt_message.insert("insert", ph)
        self.txt_message.focus()

    def _on_mode_change(self):
        mode = self.sender_mode.get()
        if mode == "twilio":
            self.frame_pywhatkit_note.pack_forget()
            self.frame_twilio.pack(fill="x", padx=16, pady=(10, 0))
        else:
            self.frame_twilio.pack_forget()
            self.frame_pywhatkit_note.pack(fill="x", padx=16, pady=(8, 0))

    def _update_contact_count(self):
        n = len(self.contacts)
        self.lbl_count.config(text=f"{n} contact{'s' if n != 1 else ''}")
        self.badge_total._label.config(text=f"{n} Contact{'s' if n != 1 else ''}")

    # ── Contact management ─────────────────────────────────────────────────────
    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, c in enumerate(self.contacts, 1):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=str(i - 1),
                             values=(i, c["name"], c["phone"],
                                     c.get("status", "Pending")),
                             tags=(tag,))
        self._update_contact_count()

    def _load_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            df = pd.read_csv(path)
            df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
            if "name" not in df.columns or "phone_number" not in df.columns:
                messagebox.showerror("Invalid CSV",
                                     "CSV must contain 'name' and 'phone_number' columns.")
                return
            df = df.dropna(subset=["name", "phone_number"])
            new_contacts = [
                {"name": str(row["name"]).strip(),
                 "phone": str(row["phone_number"]).strip(),
                 "status": "Pending"}
                for _, row in df.iterrows()
            ]
            self.contacts.extend(new_contacts)
            self._refresh_tree()
            self._log(f"Loaded {len(new_contacts)} contacts from {os.path.basename(path)}", "success")
            self._set_status(f"Loaded {len(new_contacts)} contacts from CSV.")
        except Exception as e:
            messagebox.showerror("Error loading CSV", str(e))
            self._log(f"CSV load error: {e}", "error")

    def _download_template(self):
        path = filedialog.asksaveasfilename(
            title="Save CSV template",
            defaultextension=".csv",
            initialfile="contacts_template.csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if path:
            pd.DataFrame([
                {"name": "Alice Smith", "phone_number": "+919876543210"},
                {"name": "Bob Jones",   "phone_number": "+14155238886"},
            ]).to_csv(path, index=False)
            self._log(f"Template saved to {path}", "success")

    def _add_manual_contact(self):
        name  = self.ent_name.get().strip()
        phone = self.ent_phone.get().strip()
        if not name or not phone:
            messagebox.showwarning("Missing info", "Please enter both name and phone number.")
            return
        try:
            normalize_phone(phone)
        except ValueError as e:
            messagebox.showerror("Invalid phone", str(e))
            return
        self.contacts.append({"name": name, "phone": phone, "status": "Pending"})
        self._refresh_tree()
        self.ent_name.delete(0, "end")
        self.ent_phone.delete(0, "end")
        self._log(f"Added contact: {name} ({phone})", "info")

    def _clear_contacts(self):
        if self.contacts and messagebox.askyesno("Confirm", "Clear all contacts?"):
            self.contacts.clear()
            self._refresh_tree()
            self._log("All contacts cleared.", "warn")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted([int(iid) for iid in sel], reverse=True)
        for idx in idxs:
            removed = self.contacts.pop(idx)
            self._log(f"Removed: {removed['name']}", "warn")
        self._refresh_tree()

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.ctx_menu.post(event.x_root, event.y_root)

    # ── Preview ────────────────────────────────────────────────────────────────
    def _preview_message(self):
        tmpl = self.txt_message.get("1.0", "end").strip()
        if not tmpl:
            self.lbl_preview.config(text="(no message written)")
            return
        sample_name = self.contacts[0]["name"] if self.contacts else "John Doe"
        preview = format_message(tmpl, sample_name)
        self.lbl_preview.config(text=preview, fg=PALETTE["text"])

    # ── Twilio test ────────────────────────────────────────────────────────────
    def _test_twilio(self):
        try:
            from twilio.rest import Client
            client = Client(self.twilio_sid.get(), self.twilio_token.get())
            account = client.api.accounts(self.twilio_sid.get()).fetch()
            messagebox.showinfo("Twilio OK",
                                f"Connected!\nAccount: {account.friendly_name}")
            self._log("Twilio connection test passed.", "success")
        except Exception as e:
            messagebox.showerror("Twilio Error", str(e))
            self._log(f"Twilio test failed: {e}", "error")

    # ── Sending ────────────────────────────────────────────────────────────────
    def _start_sending(self, mode: str):
        if self.is_sending:
            messagebox.showinfo("Busy", "Already sending messages. Click ■ Stop to cancel.")
            return

        template = self.txt_message.get("1.0", "end").strip()
        if not template:
            messagebox.showwarning("No message", "Please write a message in the Message tab.")
            return

        if mode == "selected":
            sel = self.tree.selection()
            targets = [self.contacts[int(iid)] for iid in sel]
        else:
            targets = self.contacts

        if not targets:
            messagebox.showwarning("No contacts",
                                   "No contacts to send to. Load a CSV or add manually.")
            return

        # Build sender
        sender_type = self.sender_mode.get()
        try:
            if sender_type == "twilio":
                sender = TwilioSender(self.twilio_sid.get(),
                                      self.twilio_token.get(),
                                      self.twilio_from.get())
            else:
                sender = PyWhatKitSender()
        except Exception as e:
            messagebox.showerror("Sender error", str(e))
            return

        delay = int(self.spin_delay.get())
        self.stop_flag.clear()
        self.is_sending = True
        self._set_status(f"Sending to {len(targets)} contacts…")

        threading.Thread(
            target=self._send_worker,
            args=(targets, template, sender, delay),
            daemon=True
        ).start()

    def _send_worker(self, targets, template, sender, delay):
        total = len(targets)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        sent_ok = 0
        sent_err = 0

        for i, contact in enumerate(targets):
            if self.stop_flag.is_set():
                self._log("Sending stopped by user.", "warn")
                break

            name  = contact["name"]
            phone = contact["phone"]
            msg   = format_message(template, name)

            # Find tree iid for this contact
            iid = str(self.contacts.index(contact)) if contact in self.contacts else None

            self._log(f"Sending to {name} ({phone})…", "info")
            try:
                sender.send(phone, msg, wait=20)
                contact["status"] = "✅ Sent"
                if iid:
                    self.tree.set(iid, "status", "✅ Sent")
                self._log(f"  ✓ Sent to {name}", "success")
                sent_ok += 1
            except Exception as e:
                contact["status"] = "❌ Failed"
                if iid:
                    self.tree.set(iid, "status", "❌ Failed")
                self._log(f"  ✗ Failed ({name}): {e}", "error")
                sent_err += 1

            self.progress["value"] = i + 1
            self.lbl_progress.config(
                text=f"{i+1}/{total}"
            )

            if i < total - 1 and not self.stop_flag.is_set():
                self._log(f"  Waiting {delay}s before next…", "info")
                time.sleep(delay)

        summary = f"Done — {sent_ok} sent, {sent_err} failed out of {total}."
        self._log(summary, "success" if sent_err == 0 else "warn")
        self._set_status(summary)
        self.is_sending = False
        self.progress["value"] = 0
        self.lbl_progress.config(text="")

    def _stop_sending(self):
        if self.is_sending:
            self.stop_flag.set()
            self._set_status("Stopping… (current message will finish)")
        else:
            self._set_status("Not currently sending.")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = WhatsAppMessengerApp()
    app.mainloop()