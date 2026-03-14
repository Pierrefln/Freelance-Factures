#!/usr/bin/env python3
"""
Générateur de Factures — Dark AI Edition
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import calendar
from datetime import date
import os
import subprocess
import sys
import json
import uuid

# ============================================================
# PALETTE DARK AI
# ============================================================

BG_MAIN     = "#070c18"
BG_CARD     = "#0d1630"
BG_INPUT    = "#0f1e38"
BG_ROW1     = "#0d1630"
BG_ROW2     = "#0a1528"
ACCENT      = "#3b82f6"
ACCENT_DIM  = "#1e3a6e"
ACCENT_HOV  = "#60a5fa"
TEXT_MAIN   = "#e2e8f0"
TEXT_MED    = "#94a3b8"
TEXT_DIM    = "#475569"
BORDER      = "#1e3460"
SUCCESS     = "#059669"
SUCCESS_HOV = "#047857"
DANGER      = "#dc2626"
DANGER_HOV  = "#b91c1c"

# ============================================================
# PROJETS
# ============================================================

PROJECTS = {
    "F0084": "PCN EIC",
    "F0110": "ACCESS2EIC",
    "F0116": "Taftie 2020",
    "F0117": "InnovI",
    "F0119": "Blue Invest",
    "F0120": "Enrich in Africa",
    "F0121": "EIC Scaling-up",
    "F0122": "Projets bilatéraux",
    "F0123": "EEN TONIC 2022 - 2028",
    "F0124": "EIC ecosystem partnerships and co-investment support",
    "F0125": "ACCELERO",
    "F0126": "Guichet Unique ivoirien",
    "F0127": "EEN2EIC",
    "F0128": "EXI AF SUBSAHARIENNE",
    "F0129": "EXI MENA",
    "F0130": "Togo BPFI",
    "F0131": "InvestEU Advisory",
    "F0132": "EXI Europe",
    "F0133": "FSPI Sénégal",
    "F0134": "InvestEU Portal",
    "F0135": "HDB PFF Facility",
    "F0136": "EIC Scaling Club",
    "F0137": "atTRACTION",
    "F0138": "Greenovi",
    "F0139": "BNI Côte d'Ivoire",
    "F0140": "SEADE",
    "F0141": "START-UP RISE \u00ab Moroccan Tech Ecosystem Support Program \u00bb (MTESP)",
    "F0142": "ESIL 2",
    "F0143": "NCC FR",
    "F0144": "Estonie EBIA TSI",
    "F0145": "ECLIPSE",
    "F0146": "EIC Access+",
    "F0149": "WE-RISE",
    "F0151": "Idealist2027",
    "F0152": "CoDEPlugin",
    "F0153": "HCDI",
    "V0147": "NCC FR - 2",
    "V0148": "Roumanie IDB",
    "V0149": "Choose Africa 2 Togo",
    "V0150": "Moldavie Digital & Economic",
    "V0151": "GUINEE FGPE",
    "V0152": "26-Kobo Art",
    "V0153": "ANGOLA \u2013 Choose Africa 2",
    "V0154": "Projets Bilatéraux Euroquity",
    "V0155": "TAFTIE",
}

MONTHS_FR = {
    1: "janvier", 2: "février", 3: "mars", 4: "avril",
    5: "mai", 6: "juin", 7: "juillet", 8: "août",
    9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre",
}

# Taux de TVA applicables en micro-entreprise
TVA_OPTIONS = [
    "Non applicable (art. 293 B CGI)",
    "20 % \u2014 Taux normal",
    "10 % \u2014 Taux intermédiaire",
    "8,5 % \u2014 Taux intermédiaire DOM",
    "5,5 % \u2014 Taux réduit",
    "2,1 % \u2014 Taux super réduit",
]

TVA_RATES = {
    "Non applicable (art. 293 B CGI)": None,
    "20 % \u2014 Taux normal": 20.0,
    "10 % \u2014 Taux intermédiaire": 10.0,
    "8,5 % \u2014 Taux intermédiaire DOM": 8.5,
    "5,5 % \u2014 Taux réduit": 5.5,
    "2,1 % \u2014 Taux super réduit": 2.1,
}

RECIPIENT = {
    "name":  "Bpifrance",
    "line1": "27/31 avenue du Général Leclerc",
    "line2": "94710 Maisons-Alfort Cedex",
}

PROFILES_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "profiles.json"
)

DEFAULT_DIR = os.path.expanduser("~")


# ============================================================
# PROFILS
# ============================================================

def load_profiles():
    """Charge les profils. Retourne (profiles, current_id)."""
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            profiles = data.get("profiles", [])
            current  = data.get("current", None)
            if profiles:
                return profiles, current
        except Exception:
            pass
    return [], None


def save_profiles(profiles, current_id):
    """Sauvegarde les profils dans le fichier JSON."""
    with open(PROFILES_FILE, "w", encoding="utf-8") as f:
        json.dump({"profiles": profiles, "current": current_id},
                  f, ensure_ascii=False, indent=2)


# ============================================================
# GÉNÉRATION PDF
# ============================================================

def fmt_amount(v):
    """Format a number as French-style euro amount: 1 200 €"""
    if v == int(v):
        s = f"{int(v):,}".replace(",", "\u202f")
    else:
        s = f"{v:.2f}".replace(".", ",")
    return s + "\u00a0\u20ac"


def generate_invoice_pdf(filepath, project_code, project_name, days,
                          month, year, invoice_day, profile, extra_desc=""):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                        Paragraph, Spacer)
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    except ImportError:
        raise ImportError("reportlab requis — lancez : pip install reportlab")

    tjm          = float(profile.get("tjm", 240))
    last_day     = calendar.monthrange(year, month)[1]
    total_ht     = days * tjm
    inv_number   = f"{project_code}-{month:02d}-{year}"
    inv_date_str = f"Le {invoice_day} {MONTHS_FR[month]} {year}"
    period_str   = (f"Période du 1er {MONTHS_FR[month]} {year} "
                    f"au {last_day} {MONTHS_FR[month]} {year}")

    # Description
    base_desc = f"Prestations de consultant dans le cadre du projet {project_name}."
    full_desc = (base_desc + "<br/><br/>" + extra_desc.replace("\n", "<br/>")
                 if extra_desc.strip() else base_desc)
    days_str = str(int(days)) if days == int(days) else str(days).replace(".", ",")

    # Données profil
    sender_name  = f"{profile.get('prenom', '').upper()} {profile.get('nom', '').upper()}"
    sender_addr1 = profile.get("adresse1", "")
    sender_addr2 = profile.get("adresse2", "")
    sender_tel   = profile.get("tel", "")
    sender_email = profile.get("email", "")
    sender_siret = profile.get("siret", "")
    bank_iban    = profile.get("iban", "")
    bank_bic     = profile.get("bic", "")
    tva_key      = profile.get("tva", "Non applicable (art. 293 B CGI)")
    tva_rate     = TVA_RATES.get(tva_key)
    full_name    = f"{profile.get('prenom', '')} {profile.get('nom', '')}"

    # Style helpers
    def S(name, **kw):
        d = dict(fontName="Helvetica", fontSize=9, leading=13)
        d.update(kw)
        return ParagraphStyle(name, **d)

    s_n  = S("n")
    s_r  = S("r",  alignment=TA_RIGHT)
    s_c  = S("c",  alignment=TA_CENTER)
    s_b  = S("b",  fontName="Helvetica-Bold")
    s_bc = S("bc", fontName="Helvetica-Bold", alignment=TA_CENTER)
    s_br = S("br", fontName="Helvetica-Bold", alignment=TA_RIGHT)
    s_rb = S("rb", fontName="Helvetica-Bold", textColor=colors.red)
    s_t  = S("t",  fontName="Helvetica-Bold", fontSize=11,
             alignment=TA_CENTER, leading=16)

    P = Paragraph

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=2.2*cm, rightMargin=2.2*cm,
        topMargin=1.5*cm, bottomMargin=2*cm,
    )
    W = A4[0] - 4.4*cm
    story = []

    # En-tête
    sender_p = P(
        f"<b>{sender_name}</b><br/>"
        f"{sender_addr1}<br/>{sender_addr2}<br/><br/>"
        f"Tél\u00a0: {sender_tel}<br/>"
        f"Email\u00a0: {sender_email}<br/><br/>"
        f"N°\u00a0SIRET\u00a0: {sender_siret}",
        s_n)
    recip_p = P(
        "<br/><br/><br/>"
        f"{RECIPIENT['name']}<br/>"
        f"{RECIPIENT['line1']}<br/>{RECIPIENT['line2']}",
        s_r)

    hdr = Table([[sender_p, recip_p]], colWidths=[W * 0.55, W * 0.45])
    hdr.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [hdr, Spacer(1, 0.7*cm)]

    story.append(P("<u><b>PAIEMENT A FIN DU MOIS</b></u>", s_rb))
    story.append(Spacer(1, 0.9*cm))

    dt = Table([[P("", s_n), P(inv_date_str, s_r)]],
               colWidths=[W * 0.6, W * 0.4])
    dt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [dt, Spacer(1, 0.5*cm)]

    story.append(P(f"<b>Facture n°{inv_number}</b>", s_t))
    story.append(P(period_str, s_c))
    story.append(Spacer(1, 0.6*cm))

    col_w = [W * 0.47, W * 0.17, W * 0.18, W * 0.18]
    main_tbl = Table(
        [
            [P("<b>Désignation des prestations</b>", s_bc),
             P("<b>Quantité (en jours)</b>", s_bc),
             P("<b>Prix HT (Hors Taxes)</b>", s_bc),
             P("<b>Prix total HT (Hors Taxes)</b>", s_bc)],
            [P(full_desc, s_n),
             P(days_str, s_c),
             P(fmt_amount(tjm), s_c),
             P(fmt_amount(total_ht), s_c)],
        ],
        colWidths=col_w,
        rowHeights=[None, 7.5*cm],
    )
    main_tbl.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",   (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
        ("LEFTPADDING",  (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(main_tbl)

    # Calcul TVA
    if tva_rate is not None:
        tva_amount = total_ht * tva_rate / 100
        total_ttc  = total_ht + tva_amount
        tva_label  = tva_key.split("\u2014")[0].strip()

    # Tableau de totaux
    if tva_rate is None:
        total_rows = [
            [P("<u><b>PAIEMENT A FIN DU MOIS</b></u>", s_rb),
             P(f"<b>Total (en euros HT)\u00a0\u00a0\u00a0\u00a0\u00a0"
               f"{fmt_amount(total_ht)}</b>", s_br)],
        ]
        total_style = [
            ("BOX",          (0, 0), (-1, -1), 0.5, colors.black),
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",   (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
            ("LEFTPADDING",  (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
    else:
        total_rows = [
            [P("<u><b>PAIEMENT A FIN DU MOIS</b></u>", s_rb),
             P(f"Total HT\u00a0\u00a0\u00a0\u00a0\u00a0{fmt_amount(total_ht)}", s_br)],
            [P("", s_n),
             P(f"TVA {tva_label}\u00a0: {fmt_amount(tva_amount)}", s_br)],
            [P("", s_n),
             P(f"<b>Total TTC\u00a0\u00a0\u00a0\u00a0\u00a0{fmt_amount(total_ttc)}</b>",
               s_br)],
        ]
        total_style = [
            ("BOX",          (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID",    (1, 0), (1, -1),  0.3, colors.HexColor("#cccccc")),
            ("SPAN",         (0, 0), (0, -1)),
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",   (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
            ("LEFTPADDING",  (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            # Ligne Total TTC avec fond léger
            ("BACKGROUND",   (1, 2), (1, 2),   colors.HexColor("#f0f4ff")),
        ]

    total_tbl = Table(total_rows, colWidths=[W * 0.55, W * 0.45])
    total_tbl.setStyle(TableStyle(total_style))
    story += [total_tbl, Spacer(1, 0.2*cm)]

    # Mention TVA sous le tableau
    if tva_rate is None:
        story.append(P("TVA Non applicable, art. 293 B du CGI", s_n))
        story.append(Spacer(1, 0.9*cm))
    else:
        story.append(Spacer(1, 0.9*cm))

    # IBAN / BIC
    iban_inner = Table(
        [["IBAN :", bank_iban],
         ["BIC :",  bank_bic]],
        colWidths=[1.6*cm, 8*cm],
    )
    iban_inner.setStyle(TableStyle([
        ("FONTNAME",     (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 9),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
    ]))
    iban_wrap = Table([[P("", s_n), iban_inner]], colWidths=[W * 0.35, W * 0.65])
    iban_wrap.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [iban_wrap, Spacer(1, 2*cm)]

    # Signature
    sig_path = profile.get("signature", "")
    if sig_path and os.path.isfile(sig_path):
        try:
            from reportlab.platypus import Image as RLImage
            sig_img = RLImage(sig_path, width=4*cm, height=2*cm,
                              kind="proportional")
            sig_cell = sig_img
        except Exception:
            sig_cell = P(f"<br/><br/><br/>{full_name}", s_c)
    else:
        sig_cell = P(f"<br/><br/><br/>{full_name}", s_c)

    sig = Table(
        [[P("", s_n), sig_cell]],
        colWidths=[W * 0.5, W * 0.5],
    )
    sig.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("ALIGN",  (1, 0), (1, 0),  "CENTER")]))
    story.append(sig)

    doc.build(story)


# ============================================================
# DIALOGUE PROFIL (création / édition)
# ============================================================

class ProfileDialog(tk.Toplevel):
    """Popup de création ou d'édition d'un profil."""

    def __init__(self, parent, profile=None, title_text="Nouveau profil"):
        super().__init__(parent)
        self.title(title_text)
        self._title_text = title_text
        self.configure(bg=BG_CARD)
        self.resizable(False, False)
        self.result = None
        self._profile = profile or {}
        self._build()
        self.grab_set()
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _make_field(self, parent, label, default, row):
        tk.Label(parent, text=label, bg=BG_CARD, fg=TEXT_MED,
                 font=("Arial", 9), anchor="e").grid(
            row=row, column=0, sticky="e", padx=(0, 10), pady=4)
        var = tk.StringVar(value=default)
        entry = tk.Entry(parent, textvariable=var,
                         bg=BG_INPUT, fg=TEXT_MAIN, relief=tk.FLAT, bd=6,
                         insertbackground=TEXT_MAIN,
                         highlightbackground=BORDER, highlightthickness=1,
                         font=("Arial", 10), width=34)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        return var

    def _section_label(self, parent, text, row):
        tk.Frame(parent, bg=BORDER, height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(8, 4))
        tk.Label(parent, text=text, bg=BG_CARD, fg=ACCENT_HOV,
                 font=("Arial", 8, "bold")).grid(
            row=row + 1, column=0, columnspan=2, sticky="w", pady=(0, 4))

    def _build(self):
        p = self._profile

        # Barre de titre colorée
        top = tk.Frame(self, bg=ACCENT_DIM, padx=20, pady=14)
        top.pack(fill=tk.X)
        tk.Label(top, text=self._title_text, bg=ACCENT_DIM, fg=TEXT_MAIN,
                 font=("Arial", 13, "bold")).pack(anchor="w")
        tk.Label(top, text="Renseignez les informations du prestataire",
                 bg=ACCENT_DIM, fg=TEXT_MED,
                 font=("Arial", 9)).pack(anchor="w")

        # Formulaire
        form = tk.Frame(self, bg=BG_CARD, padx=22, pady=16)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)

        r = 0

        # ── Identité ──
        tk.Label(form, text="IDENTITÉ", bg=BG_CARD, fg=ACCENT_HOV,
                 font=("Arial", 8, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(0, 4))
        r += 1

        self._prenom = self._make_field(form, "Prénom",   p.get("prenom",   ""), r); r += 1
        self._nom    = self._make_field(form, "Nom",      p.get("nom",      ""), r); r += 1
        self._addr1  = self._make_field(form, "Adresse",  p.get("adresse1", ""), r); r += 1
        self._addr2  = self._make_field(form, "CP + Ville", p.get("adresse2", ""), r); r += 1
        self._tel    = self._make_field(form, "Téléphone", p.get("tel",    ""), r); r += 1
        self._email  = self._make_field(form, "Email",    p.get("email",    ""), r); r += 1

        # ── Activité ──
        self._section_label(form, "ACTIVITÉ", r); r += 2

        self._siret = self._make_field(form, "SIRET",        p.get("siret",  ""), r); r += 1
        self._tjm   = self._make_field(form, "TJM (€/jour)", str(p.get("tjm", "")), r); r += 1

        tk.Label(form, text="TVA", bg=BG_CARD, fg=TEXT_MED,
                 font=("Arial", 9), anchor="e").grid(
            row=r, column=0, sticky="e", padx=(0, 10), pady=4)
        self._tva_var = tk.StringVar(value=p.get("tva", TVA_OPTIONS[0]))
        ttk.Combobox(form, textvariable=self._tva_var, values=TVA_OPTIONS,
                     state="readonly", width=33,
                     font=("Arial", 10)).grid(row=r, column=1, sticky="ew", pady=4)
        r += 1

        # ── Coordonnées bancaires ──
        self._section_label(form, "COORDONNÉES BANCAIRES", r); r += 2

        self._iban = self._make_field(form, "IBAN", p.get("iban", ""), r); r += 1
        self._bic  = self._make_field(form, "BIC",  p.get("bic",  ""), r); r += 1

        # ── Signature ──
        self._section_label(form, "SIGNATURE (optionnel)", r); r += 2

        tk.Label(form, text="Fichier image", bg=BG_CARD, fg=TEXT_MED,
                 font=("Arial", 9), anchor="e").grid(
            row=r, column=0, sticky="e", padx=(0, 10), pady=4)

        sig_row = tk.Frame(form, bg=BG_CARD)
        sig_row.grid(row=r, column=1, sticky="ew", pady=4)
        sig_row.columnconfigure(0, weight=1)

        self._sig_path = tk.StringVar(value=p.get("signature", ""))
        sig_entry = tk.Entry(sig_row, textvariable=self._sig_path,
                             bg=BG_INPUT, fg=TEXT_MAIN, relief=tk.FLAT, bd=6,
                             insertbackground=TEXT_MAIN,
                             highlightbackground=BORDER, highlightthickness=1,
                             font=("Arial", 9), width=24)
        sig_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        tk.Button(sig_row, text="Parcourir\u2026",
                  command=self._browse_signature,
                  bg=ACCENT_DIM, fg=TEXT_MAIN, relief=tk.FLAT,
                  font=("Arial", 9), padx=8, pady=4,
                  cursor="hand2", activebackground=ACCENT,
                  activeforeground="white").grid(row=0, column=1)

        tk.Label(form,
                 text="PNG, JPG ou GIF — apparaîtra en bas à droite de la facture",
                 bg=BG_CARD, fg=TEXT_DIM, font=("Arial", 8, "italic")).grid(
            row=r + 1, column=1, sticky="w")
        r += 2

        # Boutons
        btns = tk.Frame(self, bg=BG_CARD, padx=22, pady=14)
        btns.pack(fill=tk.X)

        tk.Button(btns, text="Annuler", command=self.destroy,
                  bg=BG_INPUT, fg=TEXT_MED, relief=tk.FLAT,
                  font=("Arial", 10), padx=16, pady=7,
                  cursor="hand2", activebackground=BG_MAIN,
                  activeforeground=TEXT_MAIN).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btns, text="  Enregistrer  ", command=self._save,
                  bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Arial", 10, "bold"), padx=16, pady=7,
                  cursor="hand2", activebackground=ACCENT_HOV,
                  activeforeground="white").pack(side=tk.RIGHT)

    def _browse_signature(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Choisir une image de signature",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                       ("Tous les fichiers", "*.*")],
        )
        if path:
            self._sig_path.set(path)

    def _save(self):
        prenom = self._prenom.get().strip()
        nom    = self._nom.get().strip()
        if not prenom or not nom:
            messagebox.showwarning("Champs requis",
                                   "Prénom et Nom sont obligatoires.", parent=self)
            return
        try:
            tjm = float(self._tjm.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Valeur invalide",
                                   "Le TJM doit être un nombre.", parent=self)
            return

        self.result = {
            "id":        self._profile.get("id") or str(uuid.uuid4()),
            "prenom":    prenom,
            "nom":       nom,
            "adresse1":  self._addr1.get().strip(),
            "adresse2":  self._addr2.get().strip(),
            "tel":       self._tel.get().strip(),
            "email":     self._email.get().strip(),
            "siret":     self._siret.get().strip(),
            "iban":      self._iban.get().strip(),
            "bic":       self._bic.get().strip(),
            "tjm":       tjm,
            "tva":       self._tva_var.get(),
            "signature": self._sig_path.get().strip(),
        }
        self.destroy()


# ============================================================
# DIALOGUE SÉLECTION DE PROFIL
# ============================================================

class ProfileSwitchDialog(tk.Toplevel):
    """Popup de sélection de profil actif."""

    def __init__(self, parent, profiles, current_id):
        super().__init__(parent)
        self.title("Choisir un profil")
        self.configure(bg=BG_CARD)
        self.resizable(False, False)
        self.result_id = None
        self._profiles   = profiles
        self._current_id = current_id
        self._selected   = tk.StringVar(value=current_id)
        self._build()
        self.grab_set()
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        # Titre
        top = tk.Frame(self, bg=ACCENT_DIM, padx=20, pady=14)
        top.pack(fill=tk.X)
        tk.Label(top, text="Choisir un profil", bg=ACCENT_DIM, fg=TEXT_MAIN,
                 font=("Arial", 13, "bold")).pack(anchor="w")

        # Liste des profils
        list_frame = tk.Frame(self, bg=BG_CARD, padx=18, pady=12)
        list_frame.pack(fill=tk.BOTH, expand=True)

        for p in self._profiles:
            row = tk.Frame(list_frame, bg=BG_CARD, pady=4)
            row.pack(fill=tk.X)

            rb = tk.Radiobutton(row, variable=self._selected, value=p["id"],
                                bg=BG_CARD, fg=TEXT_MAIN,
                                selectcolor=ACCENT_DIM,
                                activebackground=BG_CARD,
                                activeforeground=TEXT_MAIN)
            rb.pack(side=tk.LEFT)

            info = tk.Frame(row, bg=BG_CARD)
            info.pack(side=tk.LEFT, padx=8)

            name_txt = f"{p['prenom']} {p['nom']}"
            if p["id"] == self._current_id:
                name_txt += "  (actuel)"
            tk.Label(info, text=name_txt,
                     font=("Arial", 11, "bold"), fg=TEXT_MAIN,
                     bg=BG_CARD).pack(anchor="w")

            details = (f"SIRET\u00a0: {p.get('siret', '—')}   \u00b7   "
                       f"TJM\u00a0: {p.get('tjm', '—')}\u00a0€/j   \u00b7   "
                       f"{p.get('tva', '—')[:32]}")
            tk.Label(info, text=details, font=("Arial", 9),
                     fg=TEXT_MED, bg=BG_CARD).pack(anchor="w")

            tk.Frame(list_frame, bg=BORDER, height=1).pack(
                fill=tk.X, pady=(4, 0))

        # Boutons
        btns = tk.Frame(self, bg=BG_CARD, padx=18, pady=12)
        btns.pack(fill=tk.X)

        tk.Button(btns, text="Annuler", command=self.destroy,
                  bg=BG_INPUT, fg=TEXT_MED, relief=tk.FLAT,
                  font=("Arial", 10), padx=16, pady=7,
                  cursor="hand2", activebackground=BG_MAIN,
                  activeforeground=TEXT_MAIN).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btns, text="  Sélectionner  ", command=self._confirm,
                  bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Arial", 10, "bold"), padx=16, pady=7,
                  cursor="hand2", activebackground=ACCENT_HOV,
                  activeforeground="white").pack(side=tk.RIGHT)

    def _confirm(self):
        self.result_id = self._selected.get()
        self.destroy()


# ============================================================
# LIGNE DE FACTURE
# ============================================================

class ProjectRow(tk.Frame):
    """Une ligne = un projet + nombre de jours (+ description optionnelle)."""

    def __init__(self, parent, index, remove_cb, bg_color):
        super().__init__(parent, bg=bg_color, pady=6, padx=10)
        self._remove_cb = remove_cb

        project_list = sorted(PROJECTS.items())
        self._codes   = [c for c, _ in project_list]
        options       = [f"{c}  \u2013  {n}" for c, n in project_list]

        # Index
        tk.Label(self, text=f"#{index:02d}", bg=bg_color, fg=ACCENT_HOV,
                 font=("Courier", 10, "bold"), width=4).pack(side=tk.LEFT)

        # Projet
        self.project_var = tk.StringVar()
        ttk.Combobox(self, textvariable=self.project_var,
                     values=options, width=46, state="readonly",
                     font=("Arial", 10)).pack(side=tk.LEFT, padx=6)

        # Jours
        tk.Label(self, text="Jours :", bg=bg_color, fg=TEXT_MED,
                 font=("Arial", 10)).pack(side=tk.LEFT)
        self.days_var = tk.StringVar(value="5")
        ttk.Spinbox(self, from_=0.5, to=31, increment=0.5,
                    textvariable=self.days_var, width=6,
                    font=("Arial", 10)).pack(side=tk.LEFT, padx=4)

        # Toggle description
        self._desc_visible = False
        self._desc_btn = tk.Button(
            self, text="+ description", command=self._toggle_desc,
            bg=bg_color, fg=TEXT_DIM, relief=tk.FLAT,
            font=("Arial", 9, "italic"), cursor="hand2",
            activebackground=bg_color, activeforeground=ACCENT_HOV,
            bd=0)
        self._desc_btn.pack(side=tk.LEFT, padx=8)

        # Supprimer
        tk.Button(self, text="\u2715", command=lambda: remove_cb(self),
                  bg=DANGER, fg="white", relief=tk.FLAT,
                  width=3, font=("Arial", 10, "bold"),
                  cursor="hand2", activebackground=DANGER_HOV,
                  activeforeground="white").pack(side=tk.LEFT, padx=4)

        # Zone description (cachée)
        self._desc_frame = tk.Frame(self, bg=bg_color)
        tk.Label(self._desc_frame, text="Description :", bg=bg_color,
                 fg=TEXT_MED, font=("Arial", 9)).pack(anchor="w")
        self.desc_text = tk.Text(
            self._desc_frame, height=4, width=72,
            font=("Arial", 9), relief=tk.FLAT, bd=1,
            bg=BG_INPUT, fg=TEXT_MAIN, insertbackground=TEXT_MAIN,
            highlightbackground=BORDER, highlightthickness=1)
        self.desc_text.pack(fill=tk.X)

    def _toggle_desc(self):
        if self._desc_visible:
            self._desc_frame.pack_forget()
            self._desc_btn.configure(text="+ description")
        else:
            self._desc_frame.pack(fill=tk.X, padx=30, pady=(2, 4))
            self._desc_btn.configure(text="\u2013 description")
        self._desc_visible = not self._desc_visible

    def get_code(self):
        v = self.project_var.get()
        return v.split("  \u2013  ")[0].strip() if v else None

    def get_name(self):
        c = self.get_code()
        return PROJECTS.get(c, "") if c else ""

    def get_days(self):
        try:
            return float(self.days_var.get().replace(",", "."))
        except ValueError:
            return 0.0

    def get_extra_desc(self):
        return self.desc_text.get("1.0", tk.END).strip()

    def is_valid(self):
        return self.get_code() is not None and self.get_days() > 0


# ============================================================
# APPLICATION PRINCIPALE
# ============================================================

class App(tk.Tk):
    ROW_COLORS = (BG_ROW1, BG_ROW2)

    def __init__(self):
        super().__init__()
        self.title("Générateur de Factures \u2013 Bpifrance")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(860, 540)

        self._profiles, self._current_profile_id = load_profiles()
        self._rows = []
        self._output_dir = (DEFAULT_DIR if os.path.isdir(DEFAULT_DIR)
                            else os.path.expanduser("~"))

        self._setup_styles()
        self._build()

    # ── Styles ttk ──────────────────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TCombobox",
            fieldbackground=BG_INPUT, background=BG_CARD,
            foreground=TEXT_MAIN, arrowcolor=ACCENT_HOV,
            bordercolor=BORDER, selectbackground=ACCENT_DIM,
            selectforeground=TEXT_MAIN, insertcolor=TEXT_MAIN,
        )
        style.map("TCombobox",
            fieldbackground=[("readonly", BG_INPUT)],
            foreground=[("readonly", TEXT_MAIN)],
            bordercolor=[("focus", ACCENT)],
        )
        style.configure("TSpinbox",
            fieldbackground=BG_INPUT, background=BG_CARD,
            foreground=TEXT_MAIN, arrowcolor=ACCENT_HOV,
            bordercolor=BORDER, insertcolor=TEXT_MAIN,
        )
        style.map("TSpinbox", bordercolor=[("focus", ACCENT)])
        style.configure("TScrollbar",
            background=BG_CARD, troughcolor=BG_MAIN,
            arrowcolor=TEXT_MED, bordercolor=BORDER,
        )
        style.map("TScrollbar", background=[("active", ACCENT_DIM)])

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _get_active_profile(self):
        for p in self._profiles:
            if p["id"] == self._current_profile_id:
                return p
        return self._profiles[0] if self._profiles else None

    def _styled_btn(self, parent, text, cmd, bg_c, fg_c="white",
                    hover=None, **kwargs):
        btn = tk.Button(parent, text=text, command=cmd,
                        bg=bg_c, fg=fg_c, relief=tk.FLAT, bd=0,
                        cursor="hand2", activebackground=hover or bg_c,
                        activeforeground=fg_c, **kwargs)
        if hover:
            btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
            btn.bind("<Leave>", lambda e: btn.configure(bg=bg_c))
        return btn

    def _short(self, path, n=52):
        return path if len(path) <= n else "\u2026" + path[-(n - 1):]

    # ── Construction UI ──────────────────────────────────────────────────────

    def _build(self):
        # Encart profil
        self._profile_card_frame = tk.Frame(self, bg=BG_MAIN)
        self._profile_card_frame.pack(fill=tk.X)
        self._build_profile_card()

        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X, padx=14)

        # Barre date / dossier
        self._build_settings_bar()

        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X, padx=14, pady=(4, 0))

        # En-tête section factures
        sec = tk.Frame(self, bg=BG_MAIN, padx=14, pady=8)
        sec.pack(fill=tk.X)
        tk.Label(sec, text="Factures à générer",
                 bg=BG_MAIN, fg=TEXT_MAIN,
                 font=("Arial", 11, "bold")).pack(side=tk.LEFT)

        # Zone lignes scrollable
        self._build_rows_area()

        # Barre boutons bas
        self._build_bottom_bar()

        self._add_row()

    def _build_profile_card(self):
        for w in self._profile_card_frame.winfo_children():
            w.destroy()

        p = self._get_active_profile()

        card = tk.Frame(self._profile_card_frame, bg=BG_CARD, padx=18, pady=14)
        card.pack(fill=tk.X, padx=14, pady=(14, 6))

        if p is None:
            # Aucun profil — afficher un message d'invitation
            av = tk.Frame(card, bg="#3a2000", width=46, height=46)
            av.pack(side=tk.LEFT, padx=(0, 16))
            av.pack_propagate(False)
            tk.Label(av, text="!", bg="#3a2000", fg="#ffaa44",
                     font=("Arial", 18, "bold")).pack(expand=True)

            info = tk.Frame(card, bg=BG_CARD)
            info.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(info, text="Aucun profil configuré",
                     font=("Arial", 13, "bold"), fg="#ffaa44",
                     bg=BG_CARD).pack(anchor="w")
            tk.Label(info,
                     text="Créez un profil pour pouvoir générer des factures.",
                     font=("Arial", 9), fg=TEXT_MED, bg=BG_CARD).pack(anchor="w")

            btns = tk.Frame(card, bg=BG_CARD)
            btns.pack(side=tk.RIGHT, padx=(16, 0))
            self._styled_btn(btns, "+ Créer un profil",
                             self._new_profile_dialog,
                             SUCCESS, "white", hover=SUCCESS_HOV,
                             font=("Arial", 9), padx=12, pady=6
                             ).pack(side=tk.LEFT)
            return

        # Avatar (initiales)
        initials = (p.get("prenom", "?")[:1] + p.get("nom", "?")[:1]).upper()
        av = tk.Frame(card, bg=ACCENT_DIM, width=46, height=46)
        av.pack(side=tk.LEFT, padx=(0, 16))
        av.pack_propagate(False)
        tk.Label(av, text=initials, bg=ACCENT_DIM, fg=TEXT_MAIN,
                 font=("Arial", 15, "bold")).pack(expand=True)

        # Informations
        info = tk.Frame(card, bg=BG_CARD)
        info.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(info, text=f"{p.get('prenom', '')} {p.get('nom', '')}",
                 font=("Arial", 13, "bold"), fg=TEXT_MAIN,
                 bg=BG_CARD).pack(anchor="w")

        contact_parts = [x for x in [p.get("email"), p.get("tel")] if x]
        if contact_parts:
            tk.Label(info, text="  \u00b7  ".join(contact_parts),
                     font=("Arial", 9), fg=TEXT_MED, bg=BG_CARD).pack(anchor="w")

        tva_short = p.get("tva", "Non applicable (art. 293 B CGI)")
        details = (f"SIRET\u00a0: {p.get('siret', '\u2014')}   \u00b7   "
                   f"TJM\u00a0: {p.get('tjm', '\u2014')}\u00a0\u20ac/j   \u00b7   "
                   f"TVA\u00a0: {tva_short}")
        tk.Label(info, text=details, font=("Arial", 9),
                 fg=ACCENT_HOV, bg=BG_CARD).pack(anchor="w")

        # Boutons d'action
        btns = tk.Frame(card, bg=BG_CARD)
        btns.pack(side=tk.RIGHT, padx=(16, 0))

        self._styled_btn(btns, "\u270e Modifier",
                         self._edit_profile_dialog,
                         ACCENT_DIM, TEXT_MAIN, hover=ACCENT,
                         font=("Arial", 9), padx=12, pady=6
                         ).pack(side=tk.LEFT, padx=(0, 6))

        self._styled_btn(btns, "\u21c4 Changer",
                         self._switch_profile_dialog,
                         "#1a2d5a", TEXT_MAIN, hover="#243d7a",
                         font=("Arial", 9), padx=12, pady=6
                         ).pack(side=tk.LEFT, padx=(0, 6))

        self._styled_btn(btns, "+ Profil",
                         self._new_profile_dialog,
                         SUCCESS, "white", hover=SUCCESS_HOV,
                         font=("Arial", 9), padx=12, pady=6
                         ).pack(side=tk.LEFT)

    def _build_settings_bar(self):
        bar = tk.Frame(self, bg=BG_CARD, pady=10, padx=16)
        bar.pack(fill=tk.X, padx=14, pady=6)

        today = date.today()

        def lbl(text):
            return tk.Label(bar, text=text, bg=BG_CARD, fg=TEXT_MED,
                            font=("Arial", 10))

        lbl("Mois :").pack(side=tk.LEFT)
        month_names = [f"{i:02d} \u2013 {MONTHS_FR[i].capitalize()}"
                       for i in range(1, 13)]
        self._month_cb = ttk.Combobox(bar, values=month_names, state="readonly",
                                       width=15, font=("Arial", 10))
        self._month_cb.current(today.month - 1)
        self._month_cb.pack(side=tk.LEFT, padx=(4, 14))

        lbl("Année :").pack(side=tk.LEFT)
        self._year_var = tk.StringVar(value=str(today.year))
        ttk.Spinbox(bar, from_=2020, to=2035, textvariable=self._year_var,
                    width=7, font=("Arial", 10)).pack(side=tk.LEFT, padx=(4, 14))

        lbl("Jour de facturation :").pack(side=tk.LEFT)
        self._day_var = tk.StringVar(value=str(today.day))
        ttk.Spinbox(bar, from_=1, to=31, textvariable=self._day_var,
                    width=5, font=("Arial", 10)).pack(side=tk.LEFT, padx=(4, 20))

        lbl("Dossier :").pack(side=tk.LEFT)
        self._dir_lbl = tk.Label(bar, text=self._short(self._output_dir),
                                  bg=BG_CARD, fg=TEXT_MAIN, font=("Arial", 9))
        self._dir_lbl.pack(side=tk.LEFT, padx=(4, 6))
        self._styled_btn(bar, "Choisir\u2026", self._pick_dir,
                         ACCENT_DIM, TEXT_MAIN, hover=ACCENT,
                         font=("Arial", 9), padx=8, pady=3
                         ).pack(side=tk.LEFT)

    def _build_rows_area(self):
        outer = tk.Frame(self, bg=BG_MAIN, padx=14)
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, bg=BG_MAIN, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._rows_frame = tk.Frame(canvas, bg=BG_MAIN)
        self._rows_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_bottom_bar(self):
        btns = tk.Frame(self, bg=BG_MAIN, pady=10, padx=14)
        btns.pack(fill=tk.X)

        self._styled_btn(btns, "\uFF0B  Ajouter une facture", self._add_row,
                         SUCCESS, "white", hover=SUCCESS_HOV,
                         font=("Arial", 10, "bold"), padx=14, pady=8
                         ).pack(side=tk.LEFT)

        self._styled_btn(btns, "  Générer les PDF  ", self._generate,
                         ACCENT, "white", hover=ACCENT_HOV,
                         font=("Arial", 12, "bold"), padx=22, pady=8
                         ).pack(side=tk.RIGHT)

    # ── Gestion des profils ──────────────────────────────────────────────────

    def _new_profile_dialog(self):
        dlg = ProfileDialog(self, title_text="Nouveau profil")
        self.wait_window(dlg)
        if dlg.result:
            self._profiles.append(dlg.result)
            self._current_profile_id = dlg.result["id"]
            save_profiles(self._profiles, self._current_profile_id)
            self._build_profile_card()

    def _edit_profile_dialog(self):
        p = self._get_active_profile()
        dlg = ProfileDialog(self, profile=p, title_text="Modifier le profil")
        self.wait_window(dlg)
        if dlg.result:
            for i, prof in enumerate(self._profiles):
                if prof["id"] == p["id"]:
                    self._profiles[i] = dlg.result
                    break
            save_profiles(self._profiles, self._current_profile_id)
            self._build_profile_card()

    def _switch_profile_dialog(self):
        if len(self._profiles) <= 1:
            messagebox.showinfo(
                "Un seul profil",
                "Vous n'avez qu'un seul profil.\n"
                "Ajoutez-en un avec '+ Profil'.",
                parent=self)
            return
        dlg = ProfileSwitchDialog(self, self._profiles, self._current_profile_id)
        self.wait_window(dlg)
        if dlg.result_id and dlg.result_id != self._current_profile_id:
            self._current_profile_id = dlg.result_id
            save_profiles(self._profiles, self._current_profile_id)
            self._build_profile_card()

    # ── Lignes de factures ───────────────────────────────────────────────────

    def _add_row(self):
        idx = len(self._rows) + 1
        bg  = self.ROW_COLORS[(idx - 1) % 2]
        row = ProjectRow(self._rows_frame, idx, self._remove_row, bg)
        row.pack(fill=tk.X, pady=1)
        self._rows.append(row)

    def _remove_row(self, row):
        self._rows.remove(row)
        row.destroy()
        for i, r in enumerate(self._rows, 1):
            for w in r.winfo_children():
                if isinstance(w, tk.Label) and w.cget("text").startswith("#"):
                    w.configure(text=f"#{i:02d}")
                    break

    def _pick_dir(self):
        d = filedialog.askdirectory(initialdir=self._output_dir)
        if d:
            self._output_dir = d
            self._dir_lbl.configure(text=self._short(d))

    def _get_month(self):
        return int(self._month_cb.get().split(" \u2013 ")[0])

    # ── Génération PDF ───────────────────────────────────────────────────────

    def _generate(self):
        try:
            month = self._get_month()
            year  = int(self._year_var.get())
            day   = int(self._day_var.get())
        except ValueError:
            messagebox.showerror("Erreur", "Valeur de date invalide.")
            return

        valid = [r for r in self._rows if r.is_valid()]
        if not valid:
            messagebox.showwarning(
                "Aucune facture",
                "Sélectionnez au moins un projet et entrez le nombre de jours.")
            return

        try:
            import reportlab  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Module manquant",
                "La librairie 'reportlab' n'est pas installée.\n\n"
                "Ouvrez Terminal et exécutez :\n\n"
                "    pip install reportlab\n\n"
                "Puis relancez l'application.")
            return

        profile = self._get_active_profile()
        if profile is None:
            messagebox.showerror(
                "Profil requis",
                "Aucun profil configuré.\n\n"
                "Créez un profil avant de générer des factures.")
            return
        full_name = f"{profile.get('prenom', '')} {profile.get('nom', '')}"
        generated, errors = [], []

        for row in valid:
            code  = row.get_code()
            name  = row.get_name()
            days  = row.get_days()
            xdesc = row.get_extra_desc()
            fname = f"{full_name} - {code} - {month:02d} - {year}.pdf"
            fpath = os.path.join(self._output_dir, fname)
            try:
                generate_invoice_pdf(
                    filepath=fpath,
                    project_code=code,
                    project_name=name,
                    days=days,
                    month=month,
                    year=year,
                    invoice_day=day,
                    profile=profile,
                    extra_desc=xdesc,
                )
                generated.append(fname)
            except Exception as exc:
                errors.append(f"{code}: {exc}")

        if generated:
            lines = [f"\u2713  {f}" for f in generated]
            if errors:
                lines += ["", "Erreurs :"] + [f"\u2717  {e}" for e in errors]
            lines += ["", f"Dossier : {self._output_dir}"]
            if messagebox.askyesno("PDF générés",
                                    "\n".join(lines) + "\n\nOuvrir le dossier ?"):
                subprocess.run(["open", self._output_dir])
        else:
            messagebox.showerror("Erreur",
                                  "Aucun PDF généré.\n\n" + "\n".join(errors))


# ============================================================
# POINT D'ENTRÉE
# ============================================================

if __name__ == "__main__":
    if sys.version_info < (3, 8):
        print("Python 3.8+ requis.")
        sys.exit(1)

    app = App()
    app.mainloop()
