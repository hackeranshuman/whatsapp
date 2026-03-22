"""
WhatsApp & Email Bulk Messenger — Desktop Application
Sends personalised WhatsApp messages (pywhatkit / Twilio) AND emails (SMTP)
to clients loaded from CSV or entered manually.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import re
import os
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# THEME & PALETTE
# ─────────────────────────────────────────────────────────────────────────────
PALETTE = {
    "bg":           "#0F1117",
    "surface":      "#1A1D2E",
    "surface2":     "#242740",
    "accent":       "#25D366",
    "accent2":      "#128C7E",
    "email_accent": "#4A90D9",
    "email_dim":    "#2d6aab",
    "danger":       "#FF5C5C",
    "warning":      "#FFB347",
    "text":         "#E8EAF0",
    "text_dim":     "#8B8FA8",
    "border":       "#2E3150",
    "entry_bg":     "#151825",
    "success":      "#25D366",
}

FONT_TITLE   = ("Segoe UI", 20, "bold")
FONT_SECTION = ("Segoe UI", 11, "bold")
FONT_LABEL   = ("Segoe UI", 10)
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)
FONT_BTN     = ("Segoe UI", 10, "bold")

SMTP_PRESETS = {
    "Gmail":             ("smtp.gmail.com",          587),
    "Outlook/Hotmail":   ("smtp-mail.outlook.com",   587),
    "Yahoo Mail":        ("smtp.mail.yahoo.com",     587),
    "Zoho Mail":         ("smtp.zoho.com",           587),
    "Custom":            ("",                        587),
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def normalize_phone(phone):
    digits = re.sub(r"[^\d+]", "", str(phone).strip()).lstrip("+")
    if len(digits) < 7:
        raise ValueError(f"Phone number too short: {phone!r}")
    return digits


def is_valid_email(email):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(email).strip()))


def format_message(template, name):
    return template.replace("{name}", name)


# ─────────────────────────────────────────────────────────────────────────────
# SENDER BACKENDS
# ─────────────────────────────────────────────────────────────────────────────
class PyWhatKitSender:
    def send(self, phone, message, wait=20):
        import pywhatkit as kit
        phone_e164 = "+" + normalize_phone(phone)
        now = datetime.now()
        send_min  = now.minute + 1
        send_hour = now.hour
        if send_min >= 60:
            send_min -= 60
            send_hour = (send_hour + 1) % 24
        kit.sendwhatmsg(phone_e164, message, send_hour, send_min,
                        wait_time=wait, tab_close=True, close_time=3)


class TwilioSender:
    def __init__(self, account_sid, auth_token, from_number):
        from twilio.rest import Client
        self.client = Client(account_sid, auth_token)
        self.from_number = (from_number if from_number.startswith("whatsapp:")
                            else f"whatsapp:{from_number}")

    def send(self, phone, message, **_):
        self.client.messages.create(
            body=message,
            from_=self.from_number,
            to="whatsapp:+" + normalize_phone(phone),
        )


class EmailSender:
    def __init__(self, smtp_host, smtp_port, username, password, sender_name=""):
        self.smtp_host   = smtp_host
        self.smtp_port   = smtp_port
        self.username    = username
        self.password    = password
        self.sender_name = sender_name or username

    def test_connection(self):
        context = ssl.create_default_context()
        with smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=10) as server:
            server.ehlo()
            server.starttls(context=context)
            server.login(self.username, self.password)
        return "OK"

    def send(self, to_email, subject, body_plain, body_html=""):
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"{self.sender_name} <{self.username}>"
        msg["To"]      = to_email
        msg.attach(MIMEText(body_plain, "plain", "utf-8"))
        if body_html:
            msg.attach(MIMEText(body_html, "html", "utf-8"))
        context = ssl.create_default_context()
        with smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=15) as server:
            server.ehlo()
            server.starttls(context=context)
            server.login(self.username, self.password)
            server.sendmail(self.username, to_email, msg.as_string())


# ─────────────────────────────────────────────────────────────────────────────
# STYLED WIDGETS
# ─────────────────────────────────────────────────────────────────────────────
class StyledFrame(tk.Frame):
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PALETTE["surface"])
        kw.setdefault("relief", "flat")
        super().__init__(parent, **kw)


class SectionLabel(tk.Label):
    def __init__(self, parent, text, color=None, **kw):
        kw.setdefault("font", FONT_SECTION)
        kw.setdefault("fg", color or PALETTE["accent"])
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
    def __init__(self, parent, color=None, **kw):
        c = color or PALETTE["accent"]
        kw.setdefault("bg", c)
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
        self.title("WhatsApp & Email Bulk Messenger")
        self.geometry("1180x860")
        self.minsize(980, 700)
        self.configure(bg=PALETTE["bg"])

        self.contacts     = []
        self.sender_mode  = tk.StringVar(value="pywhatkit")
        self.is_sending   = False
        self.is_emailing  = False
        self.stop_flag    = threading.Event()
        self.email_stop   = threading.Event()

        self.twilio_sid   = tk.StringVar()
        self.twilio_token = tk.StringVar()
        self.twilio_from  = tk.StringVar()

        self.email_preset      = tk.StringVar(value="Gmail")
        self.email_smtp_host   = tk.StringVar(value="smtp.gmail.com")
        self.email_smtp_port   = tk.StringVar(value="587")
        self.email_username    = tk.StringVar()
        self.email_password    = tk.StringVar()
        self.email_sender_name = tk.StringVar()
        self.email_subject     = tk.StringVar(value="Message for {name}")
        self.email_html_mode   = tk.BooleanVar(value=False)

        self._build_header()
        self._build_main()
        self._build_status_bar()
        self._apply_ttk_styles()

    # ── TTK styles ─────────────────────────────────────────────────────────────
    def _apply_ttk_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background=PALETTE["entry_bg"], foreground=PALETTE["text"],
                        rowheight=26, fieldbackground=PALETTE["entry_bg"],
                        bordercolor=PALETTE["border"], font=FONT_LABEL)
        style.configure("Treeview.Heading",
                        background=PALETTE["surface2"], foreground=PALETTE["accent"],
                        font=FONT_SECTION, relief="flat")
        style.map("Treeview",
                  background=[("selected", PALETTE["accent2"])],
                  foreground=[("selected", "#fff")])
        style.configure("TNotebook", background=PALETTE["bg"], borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=PALETTE["surface"], foreground=PALETTE["text_dim"],
                        padding=(14, 8), font=FONT_SECTION)
        style.map("TNotebook.Tab",
                  background=[("selected", PALETTE["surface2"])],
                  foreground=[("selected", PALETTE["accent"])])
        style.configure("TProgressbar",
                        troughcolor=PALETTE["surface2"], background=PALETTE["accent"],
                        borderwidth=0, thickness=6)
        style.configure("Email.Horizontal.TProgressbar",
                        troughcolor=PALETTE["surface2"],
                        background=PALETTE["email_accent"],
                        borderwidth=0, thickness=6)
        for orient in ("Vertical", "Horizontal"):
            style.configure(f"{orient}.TScrollbar",
                            background=PALETTE["surface2"],
                            troughcolor=PALETTE["surface"],
                            borderwidth=0, arrowcolor=PALETTE["text_dim"])

    # ── Header ─────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=PALETTE["surface"])
        hdr.pack(fill="x", side="top")
        tk.Frame(hdr, bg=PALETTE["accent"], width=5).pack(side="left", fill="y")
        inner = tk.Frame(hdr, bg=PALETTE["surface"], padx=20, pady=12)
        inner.pack(side="left", fill="both", expand=True)
        tk.Label(inner, text="📱  WhatsApp & Email Bulk Messenger",
                 font=FONT_TITLE, fg=PALETTE["text"],
                 bg=PALETTE["surface"]).pack(side="left")
        tk.Label(inner, text="Send personalised messages across channels",
                 font=FONT_SMALL, fg=PALETTE["text_dim"],
                 bg=PALETTE["surface"]).pack(side="left", padx=(16, 0), pady=(5, 0))
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
        left = StyledFrame(main, width=440)
        left.pack(side="left", fill="y", padx=(0, 8))
        left.pack_propagate(False)
        self._build_left_panel(left)
        right = tk.Frame(main, bg=PALETTE["bg"])
        right.pack(side="left", fill="both", expand=True)
        self._build_right_panel(right)

    # ── Left panel ─────────────────────────────────────────────────────────────
    def _build_left_panel(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=2, pady=2)
        t_contacts = StyledFrame(nb)
        t_wa_msg   = StyledFrame(nb)
        t_email    = StyledFrame(nb)
        t_settings = StyledFrame(nb)
        nb.add(t_contacts, text="  Contacts  ")
        nb.add(t_wa_msg,   text="  WhatsApp  ")
        nb.add(t_email,    text="  Email     ")
        nb.add(t_settings, text="  Settings  ")
        self._build_contacts_tab(t_contacts)
        self._build_wa_message_tab(t_wa_msg)
        self._build_email_tab(t_email)
        self._build_settings_tab(t_settings)

    # ── Contacts tab ───────────────────────────────────────────────────────────
    def _build_contacts_tab(self, parent):
        SectionLabel(parent, "Import from CSV").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)
        tk.Label(parent,
                 text="CSV columns: name, phone_number, email  (email is optional)",
                 font=FONT_SMALL, fg=PALETTE["text_dim"], bg=PALETTE["surface"],
                 wraplength=380, justify="left").pack(anchor="w", padx=16, pady=(4, 8))
        btn_row = tk.Frame(parent, bg=PALETTE["surface"])
        btn_row.pack(fill="x", padx=16, pady=(0, 10))
        AccentButton(btn_row, text="Browse CSV",
                     command=self._load_csv).pack(side="left")
        GhostButton(btn_row, text="Download Template",
                    command=self._download_template).pack(side="left", padx=(8, 0))

        SectionLabel(parent, "Manual Entry").pack(anchor="w", padx=16, pady=(10, 2))
        self._divider(parent)
        form = tk.Frame(parent, bg=PALETTE["surface"])
        form.pack(fill="x", padx=16, pady=8)
        form.columnconfigure(0, weight=1)
        for row_idx, (label, attr) in enumerate([
            ("Name",                         "ent_name"),
            ("Phone  (e.g. +91xxxxxxxxxx)",  "ent_phone"),
            ("Email  (optional)",             "ent_email"),
        ]):
            tk.Label(form, text=label, font=FONT_SMALL,
                     fg=PALETTE["text_dim"], bg=PALETTE["surface"]
                     ).grid(row=row_idx*2, column=0, sticky="w", pady=(6, 0))
            ent = StyledEntry(form, width=30)
            ent.grid(row=row_idx*2+1, column=0, sticky="ew", pady=(2, 0))
            setattr(self, attr, ent)

        add_row = tk.Frame(parent, bg=PALETTE["surface"])
        add_row.pack(fill="x", padx=16, pady=(10, 0))
        AccentButton(add_row, text="+ Add Contact",
                     command=self._add_manual_contact).pack(side="left")
        DangerButton(add_row, text="X Clear All",
                     command=self._clear_contacts).pack(side="left", padx=(8, 0))

    # ── WhatsApp message tab ───────────────────────────────────────────────────
    def _build_wa_message_tab(self, parent):
        SectionLabel(parent, "WhatsApp Message").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)
        tk.Label(parent, text="Use {name} as a placeholder for the recipient's name.",
                 font=FONT_SMALL, fg=PALETTE["text_dim"], bg=PALETTE["surface"],
                 wraplength=380, justify="left").pack(anchor="w", padx=16, pady=(4, 6))
        qi = tk.Frame(parent, bg=PALETTE["surface"])
        qi.pack(fill="x", padx=16, pady=(0, 6))
        tk.Label(qi, text="Insert:", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]).pack(side="left")
        GhostButton(qi, text="{name}", padx=8, pady=3,
                    command=lambda: self._insert_ph(self.txt_wa_message, "{name}")
                    ).pack(side="left", padx=(6, 0))
        self.txt_wa_message = scrolledtext.ScrolledText(
            parent, height=9, bg=PALETTE["entry_bg"], fg=PALETTE["text"],
            insertbackground=PALETTE["accent"], relief="flat", font=FONT_LABEL,
            wrap="word", highlightthickness=1,
            highlightcolor=PALETTE["accent"], highlightbackground=PALETTE["border"])
        self.txt_wa_message.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        SectionLabel(parent, "Preview").pack(anchor="w", padx=16, pady=(4, 2))
        self._divider(parent)
        GhostButton(parent, text="Preview for first contact",
                    command=lambda: self._preview(self.txt_wa_message, self.lbl_wa_preview)
                    ).pack(anchor="w", padx=16, pady=6)
        self.lbl_wa_preview = tk.Label(parent, text="(preview will appear here)",
                                       font=FONT_SMALL, fg=PALETTE["text_dim"],
                                       bg=PALETTE["surface2"], wraplength=370,
                                       justify="left", padx=12, pady=8, anchor="w")
        self.lbl_wa_preview.pack(fill="x", padx=16, pady=(0, 8))

        SectionLabel(parent, "Delay Between Messages").pack(anchor="w", padx=16, pady=(4, 2))
        self._divider(parent)
        delay_row = tk.Frame(parent, bg=PALETTE["surface"])
        delay_row.pack(fill="x", padx=16, pady=8)
        tk.Label(delay_row, text="Delay (seconds):", font=FONT_LABEL,
                 fg=PALETTE["text"], bg=PALETTE["surface"]).pack(side="left")
        self.spin_wa_delay = tk.Spinbox(delay_row, from_=5, to=120, width=5,
                                        bg=PALETTE["entry_bg"], fg=PALETTE["text"],
                                        buttonbackground=PALETTE["surface2"],
                                        relief="flat", font=FONT_LABEL,
                                        highlightthickness=1,
                                        highlightbackground=PALETTE["border"])
        self.spin_wa_delay.pack(side="left", padx=(8, 0))
        self.spin_wa_delay.delete(0, "end")
        self.spin_wa_delay.insert(0, "15")

    # ── Email tab ──────────────────────────────────────────────────────────────
    def _build_email_tab(self, parent):
        SectionLabel(parent, "Email Compose",
                     color=PALETTE["email_accent"]).pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)

        subj_row = tk.Frame(parent, bg=PALETTE["surface"])
        subj_row.pack(fill="x", padx=16, pady=(8, 4))
        tk.Label(subj_row, text="Subject:", font=FONT_LABEL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"],
                 width=8, anchor="w").pack(side="left")
        StyledEntry(subj_row, textvariable=self.email_subject,
                    highlightcolor=PALETTE["email_accent"]
                    ).pack(side="left", fill="x", expand=True)

        mode_row = tk.Frame(parent, bg=PALETTE["surface"])
        mode_row.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(mode_row, text="Body mode:", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]).pack(side="left")
        for label, val in [("Plain Text", False), ("HTML", True)]:
            tk.Radiobutton(mode_row, text=label, variable=self.email_html_mode,
                           value=val, bg=PALETTE["surface"], fg=PALETTE["text"],
                           selectcolor=PALETTE["surface2"],
                           activebackground=PALETTE["surface"],
                           font=FONT_SMALL).pack(side="left", padx=(10, 0))

        qi = tk.Frame(parent, bg=PALETTE["surface"])
        qi.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(qi, text="Insert:", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]).pack(side="left")
        for ph in ["{name}", "{email}"]:
            GhostButton(qi, text=ph, padx=8, pady=3,
                        command=lambda p=ph: self._insert_ph(self.txt_email_body, p)
                        ).pack(side="left", padx=(6, 0))

        self.txt_email_body = scrolledtext.ScrolledText(
            parent, height=9, bg=PALETTE["entry_bg"], fg=PALETTE["text"],
            insertbackground=PALETTE["email_accent"], relief="flat", font=FONT_LABEL,
            wrap="word", highlightthickness=1,
            highlightcolor=PALETTE["email_accent"],
            highlightbackground=PALETTE["border"])
        self.txt_email_body.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        SectionLabel(parent, "Preview",
                     color=PALETTE["email_accent"]).pack(anchor="w", padx=16, pady=(4, 2))
        self._divider(parent)
        GhostButton(parent, text="Preview for first contact",
                    command=lambda: self._preview(self.txt_email_body,
                                                  self.lbl_email_preview, subject=True)
                    ).pack(anchor="w", padx=16, pady=6)
        self.lbl_email_preview = tk.Label(
            parent, text="(preview will appear here)",
            font=FONT_SMALL, fg=PALETTE["text_dim"],
            bg=PALETTE["surface2"], wraplength=370,
            justify="left", padx=12, pady=8, anchor="w")
        self.lbl_email_preview.pack(fill="x", padx=16, pady=(0, 8))

        SectionLabel(parent, "Delay Between Emails",
                     color=PALETTE["email_accent"]).pack(anchor="w", padx=16, pady=(4, 2))
        self._divider(parent)
        delay_row = tk.Frame(parent, bg=PALETTE["surface"])
        delay_row.pack(fill="x", padx=16, pady=8)
        tk.Label(delay_row, text="Delay (seconds):", font=FONT_LABEL,
                 fg=PALETTE["text"], bg=PALETTE["surface"]).pack(side="left")
        self.spin_email_delay = tk.Spinbox(delay_row, from_=1, to=60, width=5,
                                           bg=PALETTE["entry_bg"], fg=PALETTE["text"],
                                           buttonbackground=PALETTE["surface2"],
                                           relief="flat", font=FONT_LABEL,
                                           highlightthickness=1,
                                           highlightbackground=PALETTE["border"])
        self.spin_email_delay.pack(side="left", padx=(8, 0))
        self.spin_email_delay.delete(0, "end")
        self.spin_email_delay.insert(0, "3")

    # ── Settings tab ───────────────────────────────────────────────────────────
    def _build_settings_tab(self, parent):
        SectionLabel(parent, "WhatsApp Method").pack(anchor="w", padx=16, pady=(14, 2))
        self._divider(parent)
        for val, label in [
            ("pywhatkit", "PyWhatKit  (WhatsApp Web - no API key)"),
            ("twilio",    "Twilio API  (production)"),
        ]:
            tk.Radiobutton(parent, text=label, variable=self.sender_mode, value=val,
                           bg=PALETTE["surface"], fg=PALETTE["text"],
                           selectcolor=PALETTE["surface2"],
                           activebackground=PALETTE["surface"],
                           activeforeground=PALETTE["accent"],
                           font=FONT_LABEL, anchor="w",
                           command=self._on_wa_mode_change
                           ).pack(fill="x", padx=20, pady=3)

        self.frame_pywhatkit_note = tk.Frame(parent, bg=PALETTE["surface2"], padx=14, pady=8)
        self.frame_pywhatkit_note.pack(fill="x", padx=16, pady=(6, 0))
        tk.Label(self.frame_pywhatkit_note,
                 text="Open WhatsApp Web in Chrome before sending.",
                 font=FONT_SMALL, fg=PALETTE["warning"],
                 bg=PALETTE["surface2"], justify="left").pack(anchor="w")

        self.frame_twilio = tk.Frame(parent, bg=PALETTE["surface"])
        self.frame_twilio.pack(fill="x", padx=16, pady=(8, 0))
        SectionLabel(self.frame_twilio, "Twilio Credentials").pack(anchor="w", pady=(4, 2))
        self._divider(self.frame_twilio)
        for label, var, secret in [
            ("Account SID",   self.twilio_sid,   False),
            ("Auth Token",    self.twilio_token, True),
            ("From Number",   self.twilio_from,  False),
        ]:
            tk.Label(self.frame_twilio, text=label, font=FONT_SMALL,
                     fg=PALETTE["text_dim"], bg=PALETTE["surface"]
                     ).pack(anchor="w", pady=(5, 0))
            StyledEntry(self.frame_twilio, textvariable=var,
                        show="*" if secret else "").pack(fill="x", pady=(2, 0))
        AccentButton(self.frame_twilio, text="Test Twilio",
                     command=self._test_twilio).pack(pady=8)
        self.frame_twilio.pack_forget()

        # ── Email SMTP ─────────────────────────────────────────────────────────
        SectionLabel(parent, "Email / SMTP Settings",
                     color=PALETTE["email_accent"]).pack(anchor="w", padx=16, pady=(16, 2))
        self._divider(parent)

        smtp_frame = tk.Frame(parent, bg=PALETTE["surface"])
        smtp_frame.pack(fill="x", padx=16, pady=8)
        smtp_frame.columnconfigure(1, weight=1)

        tk.Label(smtp_frame, text="Provider:", font=FONT_SMALL,
                 fg=PALETTE["text_dim"], bg=PALETTE["surface"]
                 ).grid(row=0, column=0, sticky="w", pady=3)
        preset_menu = ttk.Combobox(smtp_frame, textvariable=self.email_preset,
                                   values=list(SMTP_PRESETS.keys()),
                                   state="readonly", width=22)
        preset_menu.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=3)
        preset_menu.bind("<<ComboboxSelected>>", self._on_email_preset)

        for row, (label, var, secret) in enumerate([
            ("SMTP Host:",               self.email_smtp_host,   False),
            ("SMTP Port:",               self.email_smtp_port,   False),
            ("Your Email:",              self.email_username,    False),
            ("Password / App Password:", self.email_password,    True),
            ("Sender Name (optional):",  self.email_sender_name, False),
        ], start=1):
            tk.Label(smtp_frame, text=label, font=FONT_SMALL,
                     fg=PALETTE["text_dim"], bg=PALETTE["surface"]
                     ).grid(row=row, column=0, sticky="w", pady=3)
            StyledEntry(smtp_frame, textvariable=var,
                        show="*" if secret else "",
                        highlightcolor=PALETTE["email_accent"]
                        ).grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=3)

        tk.Label(smtp_frame,
                 text="Gmail: enable 2FA then use an App Password\n"
                      "(myaccount.google.com > Security > App Passwords)",
                 font=FONT_SMALL, fg=PALETTE["warning"],
                 bg=PALETTE["surface2"], justify="left",
                 padx=8, pady=6, wraplength=340
                 ).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        AccentButton(smtp_frame, text="Test Email Connection",
                     color=PALETTE["email_accent"],
                     command=self._test_email
                     ).grid(row=7, column=0, columnspan=2, pady=10, sticky="w")

    # ── Right panel ────────────────────────────────────────────────────────────
    def _build_right_panel(self, parent):
        top = StyledFrame(parent)
        top.pack(fill="both", expand=True, pady=(0, 8))

        hdr = tk.Frame(top, bg=PALETTE["surface"])
        hdr.pack(fill="x", padx=16, pady=(14, 4))
        SectionLabel(hdr, "Contact List").pack(side="left")
        self.lbl_count = tk.Label(hdr, text="0 contacts",
                                  font=FONT_SMALL, fg=PALETTE["text_dim"],
                                  bg=PALETTE["surface"])
        self.lbl_count.pack(side="left", padx=10)
        self._divider(top)

        tree_frame = tk.Frame(top, bg=PALETTE["surface"])
        tree_frame.pack(fill="both", expand=True, padx=16, pady=8)

        cols = ("no", "name", "phone", "email", "wa_status", "email_status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
        for col, w, heading in [
            ("no",           40,  "#"),
            ("name",        150,  "Name"),
            ("phone",       140,  "Phone"),
            ("email",       190,  "Email"),
            ("wa_status",   100,  "WhatsApp"),
            ("email_status",100,  "Email"),
        ]:
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=w, minwidth=30,
                             anchor="center" if col == "no" else "w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.ctx_menu = tk.Menu(self, tearoff=0, bg=PALETTE["surface2"],
                                fg=PALETTE["text"],
                                activebackground=PALETTE["accent2"],
                                font=FONT_SMALL)
        self.ctx_menu.add_command(label="Remove selected", command=self._remove_selected)
        self.tree.bind("<Button-3>", self._show_ctx_menu)

        # WhatsApp send bar
        wa_bar = tk.Frame(top, bg=PALETTE["surface"])
        wa_bar.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(wa_bar, text="WhatsApp:", font=FONT_SMALL,
                 fg=PALETTE["accent"], bg=PALETTE["surface"],
                 width=11, anchor="w").pack(side="left")
        AccentButton(wa_bar, text="Send All",
                     command=lambda: self._start_wa("all")).pack(side="left")
        GhostButton(wa_bar, text="Selected",
                    command=lambda: self._start_wa("selected")).pack(side="left", padx=(6, 0))
        DangerButton(wa_bar, text="Stop",
                     command=self._stop_wa).pack(side="left", padx=(6, 0))
        self.lbl_wa_prog = tk.Label(wa_bar, text="", font=FONT_SMALL,
                                    fg=PALETTE["text_dim"], bg=PALETTE["surface"])
        self.lbl_wa_prog.pack(side="right", padx=(0, 8))
        self.progress_wa = ttk.Progressbar(wa_bar, orient="horizontal",
                                           mode="determinate", length=160)
        self.progress_wa.pack(side="right")

        # Email send bar
        em_bar = tk.Frame(top, bg=PALETTE["surface"])
        em_bar.pack(fill="x", padx=16, pady=(0, 14))
        tk.Label(em_bar, text="Email:", font=FONT_SMALL,
                 fg=PALETTE["email_accent"], bg=PALETTE["surface"],
                 width=11, anchor="w").pack(side="left")
        AccentButton(em_bar, text="Send All", color=PALETTE["email_accent"],
                     command=lambda: self._start_email("all")).pack(side="left")
        GhostButton(em_bar, text="Selected",
                    command=lambda: self._start_email("selected")).pack(side="left", padx=(6, 0))
        DangerButton(em_bar, text="Stop",
                     command=self._stop_email).pack(side="left", padx=(6, 0))
        self.lbl_em_prog = tk.Label(em_bar, text="", font=FONT_SMALL,
                                    fg=PALETTE["text_dim"], bg=PALETTE["surface"])
        self.lbl_em_prog.pack(side="right", padx=(0, 8))
        self.progress_email = ttk.Progressbar(
            em_bar, orient="horizontal",
            style="Email.Horizontal.TProgressbar",
            mode="determinate", length=160)
        self.progress_email.pack(side="right")

        # Log
        log_frame = StyledFrame(parent)
        log_frame.pack(fill="both", expand=False)
        log_hdr = tk.Frame(log_frame, bg=PALETTE["surface"])
        log_hdr.pack(fill="x", padx=16, pady=(10, 4))
        SectionLabel(log_hdr, "Activity Log").pack(side="left")
        GhostButton(log_hdr, text="Clear Log", padx=8, pady=3,
                    command=self._clear_log).pack(side="right")
        self._divider(log_frame)
        self.txt_log = scrolledtext.ScrolledText(
            log_frame, height=7, bg=PALETTE["entry_bg"], fg=PALETTE["text"],
            font=FONT_MONO, state="disabled", relief="flat", highlightthickness=0)
        self.txt_log.pack(fill="both", expand=True, padx=16, pady=(6, 12))
        self.txt_log.tag_config("success", foreground=PALETTE["success"])
        self.txt_log.tag_config("error",   foreground=PALETTE["danger"])
        self.txt_log.tag_config("info",    foreground=PALETTE["text_dim"])
        self.txt_log.tag_config("warn",    foreground=PALETTE["warning"])
        self.txt_log.tag_config("email",   foreground=PALETTE["email_accent"])

    # ── Status bar ─────────────────────────────────────────────────────────────
    def _build_status_bar(self):
        bar = tk.Frame(self, bg=PALETTE["surface2"], height=28)
        bar.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(
            bar, text="Ready  -  Load contacts, compose your message, then send",
            font=FONT_SMALL, fg=PALETTE["text_dim"], bg=PALETTE["surface2"])
        self.lbl_status.pack(side="left", padx=14)
        tk.Label(bar, text="v2.0  -  WhatsApp & Email Messenger",
                 font=FONT_SMALL, fg=PALETTE["border"],
                 bg=PALETTE["surface2"]).pack(side="right", padx=14)

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _divider(self, parent):
        tk.Frame(parent, bg=PALETTE["border"], height=1).pack(fill="x", padx=16)

    def _set_status(self, msg):
        self.lbl_status.config(text=msg)

    def _log(self, msg, tag="info"):
        self.txt_log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {msg}\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _clear_log(self):
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")

    def _insert_ph(self, widget, ph):
        widget.insert("insert", ph)
        widget.focus()

    def _on_wa_mode_change(self):
        if self.sender_mode.get() == "twilio":
            self.frame_pywhatkit_note.pack_forget()
            self.frame_twilio.pack(fill="x", padx=16, pady=(8, 0))
        else:
            self.frame_twilio.pack_forget()
            self.frame_pywhatkit_note.pack(fill="x", padx=16, pady=(6, 0))

    def _on_email_preset(self, _event=None):
        host, port = SMTP_PRESETS.get(self.email_preset.get(), ("", 587))
        self.email_smtp_host.set(host)
        self.email_smtp_port.set(str(port))

    def _update_count(self):
        n = len(self.contacts)
        self.lbl_count.config(text=f"{n} contact{'s' if n != 1 else ''}")
        self.badge_total._label.config(text=f"{n} Contact{'s' if n != 1 else ''}")

    # ── Contact management ─────────────────────────────────────────────────────
    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, c in enumerate(self.contacts, 1):
            self.tree.insert("", "end", iid=str(i - 1),
                             values=(i, c.get("name",""), c.get("phone",""),
                                     c.get("email",""),
                                     c.get("wa_status","Pending"),
                                     c.get("email_status","--")))
        self._update_count()

    def _load_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
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
            added = 0
            for _, row in df.iterrows():
                self.contacts.append({
                    "name":         str(row["name"]).strip(),
                    "phone":        str(row["phone_number"]).strip(),
                    "email":        str(row.get("email","")).strip() if "email" in df.columns else "",
                    "wa_status":    "Pending",
                    "email_status": "--",
                })
                added += 1
            self._refresh_tree()
            self._log(f"Loaded {added} contacts from {os.path.basename(path)}", "success")
            self._set_status(f"Loaded {added} contacts.")
        except Exception as e:
            messagebox.showerror("CSV Error", str(e))
            self._log(f"CSV load error: {e}", "error")

    def _download_template(self):
        path = filedialog.asksaveasfilename(
            title="Save template", defaultextension=".csv",
            initialfile="contacts_template.csv",
            filetypes=[("CSV files", "*.csv")])
        if path:
            pd.DataFrame([
                {"name": "Alice Smith",  "phone_number": "+919876543210", "email": "alice@example.com"},
                {"name": "Bob Jones",    "phone_number": "+14155238886",  "email": "bob@example.com"},
            ]).to_csv(path, index=False)
            self._log(f"Template saved: {path}", "success")

    def _add_manual_contact(self):
        name  = self.ent_name.get().strip()
        phone = self.ent_phone.get().strip()
        email = self.ent_email.get().strip()
        if not name or not phone:
            messagebox.showwarning("Missing info", "Name and phone are required.")
            return
        try:
            normalize_phone(phone)
        except ValueError as e:
            messagebox.showerror("Invalid phone", str(e))
            return
        if email and not is_valid_email(email):
            messagebox.showerror("Invalid email", f"'{email}' is not a valid email address.")
            return
        self.contacts.append({"name": name, "phone": phone, "email": email,
                               "wa_status": "Pending", "email_status": "--"})
        self._refresh_tree()
        for e in (self.ent_name, self.ent_phone, self.ent_email):
            e.delete(0, "end")
        self._log(f"Added: {name}  {phone}  {email}", "info")

    def _clear_contacts(self):
        if self.contacts and messagebox.askyesno("Confirm", "Clear all contacts?"):
            self.contacts.clear()
            self._refresh_tree()
            self._log("All contacts cleared.", "warn")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        for idx in sorted([int(i) for i in sel], reverse=True):
            removed = self.contacts.pop(idx)
            self._log(f"Removed: {removed['name']}", "warn")
        self._refresh_tree()

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.ctx_menu.post(event.x_root, event.y_root)

    # ── Preview ────────────────────────────────────────────────────────────────
    def _preview(self, text_widget, label_widget, subject=False):
        body = text_widget.get("1.0", "end").strip()
        if not body:
            label_widget.config(text="(no message written)")
            return
        sample = self.contacts[0] if self.contacts else {"name": "John Doe", "email": "john@example.com"}
        preview = format_message(body, sample.get("name", "John Doe"))
        preview = preview.replace("{email}", sample.get("email", ""))
        if subject:
            subj = format_message(self.email_subject.get(), sample.get("name", "John Doe"))
            preview = f"Subject: {subj}\n\n{preview}"
        label_widget.config(text=preview[:400] + ("..." if len(preview) > 400 else ""),
                            fg=PALETTE["text"])

    # ── Twilio test ────────────────────────────────────────────────────────────
    def _test_twilio(self):
        try:
            from twilio.rest import Client
            client = Client(self.twilio_sid.get(), self.twilio_token.get())
            acc = client.api.accounts(self.twilio_sid.get()).fetch()
            messagebox.showinfo("Twilio OK", f"Connected!\nAccount: {acc.friendly_name}")
            self._log("Twilio test passed.", "success")
        except Exception as e:
            messagebox.showerror("Twilio Error", str(e))
            self._log(f"Twilio test failed: {e}", "error")

    # ── Email test ─────────────────────────────────────────────────────────────
    def _test_email(self):
        try:
            sender = self._build_email_sender()
            sender.test_connection()
            messagebox.showinfo("Email OK",
                                f"Connected to {self.email_smtp_host.get()} successfully!")
            self._log("Email SMTP test passed.", "email")
        except Exception as e:
            messagebox.showerror("Email Error",
                                 f"Could not connect:\n{e}\n\n"
                                 "Gmail tip: use an App Password, not your main password.")
            self._log(f"Email test failed: {e}", "error")

    def _build_email_sender(self):
        host = self.email_smtp_host.get().strip()
        port = int(self.email_smtp_port.get().strip() or 587)
        user = self.email_username.get().strip()
        pwd  = self.email_password.get()
        name = self.email_sender_name.get().strip()
        if not host or not user or not pwd:
            raise ValueError("SMTP host, email address and password are all required.")
        return EmailSender(host, port, user, pwd, name)

    # ── WhatsApp send ──────────────────────────────────────────────────────────
    def _start_wa(self, mode):
        if self.is_sending:
            messagebox.showinfo("Busy", "Already sending WhatsApp messages.")
            return
        template = self.txt_wa_message.get("1.0", "end").strip()
        if not template:
            messagebox.showwarning("No message", "Write a WhatsApp message first.")
            return
        targets = self._get_targets(mode)
        if not targets:
            messagebox.showwarning("No contacts", "No contacts to send to.")
            return
        try:
            if self.sender_mode.get() == "twilio":
                sender = TwilioSender(self.twilio_sid.get(),
                                      self.twilio_token.get(),
                                      self.twilio_from.get())
            else:
                sender = PyWhatKitSender()
        except Exception as e:
            messagebox.showerror("Sender error", str(e))
            return
        delay = int(self.spin_wa_delay.get())
        self.stop_flag.clear()
        self.is_sending = True
        self._set_status(f"Sending WhatsApp to {len(targets)} contacts...")
        threading.Thread(target=self._wa_worker,
                         args=(targets, template, sender, delay),
                         daemon=True).start()

    def _wa_worker(self, targets, template, sender, delay):
        total = len(targets)
        self.progress_wa["maximum"] = total
        ok = err = 0
        for i, c in enumerate(targets):
            if self.stop_flag.is_set():
                self._log("WhatsApp sending stopped.", "warn")
                break
            name, phone = c["name"], c["phone"]
            self._log(f"WA -> {name} ({phone})", "info")
            try:
                sender.send(phone, format_message(template, name), wait=20)
                c["wa_status"] = "Sent"
                self._set_tree_value(c, "wa_status", "Sent")
                self._log(f"  WA sent to {name}", "success")
                ok += 1
            except Exception as e:
                c["wa_status"] = "Failed"
                self._set_tree_value(c, "wa_status", "Failed")
                self._log(f"  WA failed ({name}): {e}", "error")
                err += 1
            self.progress_wa["value"] = i + 1
            self.lbl_wa_prog.config(text=f"{i+1}/{total}")
            if i < total - 1 and not self.stop_flag.is_set():
                time.sleep(delay)
        summary = f"WhatsApp done -- {ok} sent, {err} failed."
        self._log(summary, "success" if err == 0 else "warn")
        self._set_status(summary)
        self.is_sending = False
        self.progress_wa["value"] = 0
        self.lbl_wa_prog.config(text="")

    def _stop_wa(self):
        if self.is_sending:
            self.stop_flag.set()
            self._set_status("Stopping WhatsApp...")

    # ── Email send ─────────────────────────────────────────────────────────────
    def _start_email(self, mode):
        if self.is_emailing:
            messagebox.showinfo("Busy", "Already sending emails.")
            return
        body_template = self.txt_email_body.get("1.0", "end").strip()
        subj_template = self.email_subject.get().strip()
        if not body_template:
            messagebox.showwarning("No message", "Write an email body first.")
            return
        if not subj_template:
            messagebox.showwarning("No subject", "Enter an email subject.")
            return
        targets = self._get_targets(mode, require_email=True)
        if not targets:
            messagebox.showwarning("No valid targets",
                                   "No contacts with valid email addresses found.\n"
                                   "Add email addresses to your contacts or CSV.")
            return
        try:
            sender = self._build_email_sender()
        except Exception as e:
            messagebox.showerror("Config error", str(e))
            return
        delay     = int(self.spin_email_delay.get())
        html_mode = self.email_html_mode.get()
        self.email_stop.clear()
        self.is_emailing = True
        self._set_status(f"Sending email to {len(targets)} contacts...")
        threading.Thread(target=self._email_worker,
                         args=(targets, subj_template, body_template,
                               sender, delay, html_mode),
                         daemon=True).start()

    def _email_worker(self, targets, subj_tmpl, body_tmpl, sender, delay, html_mode):
        total = len(targets)
        self.progress_email["maximum"] = total
        ok = err = 0
        for i, c in enumerate(targets):
            if self.email_stop.is_set():
                self._log("Email sending stopped.", "warn")
                break
            name  = c["name"]
            email = c["email"]
            subj  = format_message(subj_tmpl, name)
            body  = format_message(body_tmpl, name).replace("{email}", email)
            self._log(f"Email -> {name} <{email}>", "email")
            try:
                sender.send(email, subj, body, body if html_mode else "")
                c["email_status"] = "Sent"
                self._set_tree_value(c, "email_status", "Sent")
                self._log(f"  Email sent to {name}", "success")
                ok += 1
            except Exception as e:
                c["email_status"] = "Failed"
                self._set_tree_value(c, "email_status", "Failed")
                self._log(f"  Email failed ({name}): {e}", "error")
                err += 1
            self.progress_email["value"] = i + 1
            self.lbl_em_prog.config(text=f"{i+1}/{total}")
            if i < total - 1 and not self.email_stop.is_set():
                time.sleep(delay)
        summary = f"Email done -- {ok} sent, {err} failed."
        self._log(summary, "success" if err == 0 else "warn")
        self._set_status(summary)
        self.is_emailing = False
        self.progress_email["value"] = 0
        self.lbl_em_prog.config(text="")

    def _stop_email(self):
        if self.is_emailing:
            self.email_stop.set()
            self._set_status("Stopping email...")

    # ── Shared helpers ─────────────────────────────────────────────────────────
    def _get_targets(self, mode, require_email=False):
        if mode == "selected":
            sel  = self.tree.selection()
            pool = [self.contacts[int(iid)] for iid in sel]
        else:
            pool = self.contacts
        if require_email:
            valid   = [c for c in pool if c.get("email") and is_valid_email(c["email"])]
            skipped = len(pool) - len(valid)
            if skipped:
                self._log(f"Skipping {skipped} contact(s) -- no valid email.", "warn")
            return valid
        return pool

    def _set_tree_value(self, contact, column, value):
        try:
            idx = self.contacts.index(contact)
            self.tree.set(str(idx), column, value)
        except (ValueError, tk.TclError):
            pass


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = WhatsAppMessengerApp()
    app.mainloop()