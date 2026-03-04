#!/usr/bin/env python3
"""
Générateur de Factures — Pierre-Alexis Fillon / Bpifrance
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import calendar
from datetime import date
import os
import subprocess
import sys

# ============================================================
# DONNÉES FIXES
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

PRIX_HT = 240  # € par jour

SENDER = {
    "name":    "PIERRE-ALEXIS FILLON",
    "line1":   "19 rue Voltaire",
    "line2":   "94700 Maisons-Alfort",
    "tel":     "07 81 50 50 27",
    "email":   "pierre.alexisf@gmail.com",
    "siret":   "98848421700017",
}

RECIPIENT = {
    "name":  "Bpifrance",
    "line1": "27/31 avenue du Général Leclerc",
    "line2": "94710 Maisons-Alfort Cedex",
}

BANK = {
    "iban": "FR7640618804860004068190345",
    "bic":  "BOUSFRPPXXX",
}

DEFAULT_DIR = "/Users/pa-fillon/Factures BPI"


# ============================================================
# GÉNÉRATION PDF
# ============================================================

def fmt_amount(v):
    """Format a number as French-style euro amount: 1 200 €"""
    if v == int(v):
        s = f"{int(v):,}".replace(",", "\u202f")  # narrow no-break space
    else:
        s = f"{v:.2f}".replace(".", ",")
    return s + "\u00a0\u20ac"


def generate_invoice_pdf(filepath, project_code, project_name, days,
                          month, year, invoice_day, extra_desc=""):
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

    last_day     = calendar.monthrange(year, month)[1]
    total_ht     = days * PRIX_HT
    inv_number   = f"{project_code}-{month:02d}-{year}"
    inv_date_str = f"Le {invoice_day} {MONTHS_FR[month]} {year}"
    period_str   = (f"Période du 1er {MONTHS_FR[month]} {year} "
                    f"au {last_day} {MONTHS_FR[month]} {year}")

    # Description
    base_desc = f"Prestations de consultant dans le cadre du projet {project_name}."
    full_desc = (base_desc + "<br/><br/>" + extra_desc.replace("\n", "<br/>")
                 if extra_desc.strip() else base_desc)

    # Days display (French decimal)
    days_str = str(int(days)) if days == int(days) else str(days).replace(".", ",")

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
    W = A4[0] - 4.4*cm   # usable width ≈ 16.7 cm
    story = []

    # ── En-tête : expéditeur (gauche) / destinataire (droite) ──────────────
    sender_p = P(
        f"<b>{SENDER['name']}</b><br/>"
        f"{SENDER['line1']}<br/>{SENDER['line2']}<br/><br/>"
        f"Tél\u00a0: {SENDER['tel']}<br/>"
        f"Email\u00a0: {SENDER['email']}<br/><br/>"
        f"N°\u00a0SIRET\u00a0: {SENDER['siret']}",
        s_n)
    recip_p = P(
        "<br/><br/><br/>"
        f"{RECIPIENT['name']}<br/>"
        f"{RECIPIENT['line1']}<br/>{RECIPIENT['line2']}",
        s_r)

    hdr = Table([[sender_p, recip_p]], colWidths=[W * 0.55, W * 0.45])
    hdr.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [hdr, Spacer(1, 0.7*cm)]

    # ── PAIEMENT A FIN DU MOIS ─────────────────────────────────────────────
    story.append(P("<u><b>PAIEMENT A FIN DU MOIS</b></u>", s_rb))
    story.append(Spacer(1, 0.9*cm))

    # ── Date (droite) ──────────────────────────────────────────────────────
    dt = Table([[P("", s_n), P(inv_date_str, s_r)]],
               colWidths=[W * 0.6, W * 0.4])
    dt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [dt, Spacer(1, 0.5*cm)]

    # ── Titre de la facture ────────────────────────────────────────────────
    story.append(P(f"<b>Facture n°{inv_number}</b>", s_t))
    story.append(P(period_str, s_c))
    story.append(Spacer(1, 0.6*cm))

    # ── Tableau principal ──────────────────────────────────────────────────
    col_w = [W * 0.47, W * 0.17, W * 0.18, W * 0.18]
    main_tbl = Table(
        [
            [P("<b>Désignation des prestations</b>", s_bc),
             P("<b>Quantité (en jours)</b>", s_bc),
             P("<b>Prix HT (Hors Taxes)</b>", s_bc),
             P("<b>Prix total HT (Hors Taxes)</b>", s_bc)],
            [P(full_desc, s_n),
             P(days_str, s_c),
             P(fmt_amount(PRIX_HT), s_c),
             P(fmt_amount(total_ht), s_c)],
        ],
        colWidths=col_w,
        rowHeights=[None, 7.5*cm],
    )
    main_tbl.setStyle(TableStyle([
        ("BOX",         (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID",   (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",  (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",(0, 0), (-1, -1), 5),
    ]))
    story.append(main_tbl)

    # ── Ligne Total ────────────────────────────────────────────────────────
    total_tbl = Table(
        [[P("<u><b>PAIEMENT A FIN DU MOIS</b></u>", s_rb),
          P(f"<b>Total (en euros HT)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            f"{fmt_amount(total_ht)}</b>", s_br)]],
        colWidths=[W * 0.55, W * 0.45],
    )
    total_tbl.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story += [total_tbl, Spacer(1, 0.2*cm)]

    # ── TVA ────────────────────────────────────────────────────────────────
    story.append(P("TVA Non applicable, art. 293 B du CGI", s_n))
    story.append(Spacer(1, 0.9*cm))

    # ── IBAN / BIC (aligné à droite) ───────────────────────────────────────
    iban_inner = Table(
        [["IBAN :", BANK["iban"]],
         ["BIC :",  BANK["bic"]]],
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

    # ── Signature ──────────────────────────────────────────────────────────
    sig = Table(
        [[P("", s_n), P("<br/><br/><br/>Pierre-Alexis Fillon", s_c)]],
        colWidths=[W * 0.5, W * 0.5],
    )
    sig.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(sig)

    doc.build(story)


# ============================================================
# INTERFACE GRAPHIQUE
# ============================================================

class ProjectRow(tk.Frame):
    """Une ligne = un projet + nombre de jours (+ description optionnelle)."""

    def __init__(self, parent, index, remove_cb, bg_color):
        super().__init__(parent, bg=bg_color, pady=4, padx=6)
        self._remove_cb = remove_cb

        project_list = sorted(PROJECTS.items())
        self._codes   = [c for c, _ in project_list]
        options       = [f"{c}  –  {n}" for c, n in project_list]

        # Index label
        tk.Label(self, text=f"#{index}", bg=bg_color, fg="#888",
                 font=("Arial", 10), width=3).pack(side=tk.LEFT)

        # Project combo
        self.project_var = tk.StringVar()
        ttk.Combobox(self, textvariable=self.project_var,
                     values=options, width=46, state="readonly",
                     font=("Arial", 10)).pack(side=tk.LEFT, padx=6)

        # Days
        tk.Label(self, text="Jours :", bg=bg_color,
                 font=("Arial", 10)).pack(side=tk.LEFT)
        self.days_var = tk.StringVar(value="5")
        ttk.Spinbox(self, from_=0.5, to=31, increment=0.5,
                    textvariable=self.days_var, width=6,
                    font=("Arial", 10)).pack(side=tk.LEFT, padx=4)

        # Description toggle
        self._desc_visible = False
        self._desc_btn = tk.Button(
            self, text="+ description", command=self._toggle_desc,
            bg=bg_color, fg="#666", relief=tk.FLAT,
            font=("Arial", 9, "italic"))
        self._desc_btn.pack(side=tk.LEFT, padx=8)

        # Remove
        tk.Button(self, text="✕", command=lambda: remove_cb(self),
                  bg="#ff6b6b", fg="white", relief=tk.FLAT,
                  width=3, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=4)

        # Hidden description area
        self._desc_frame = tk.Frame(self, bg=bg_color)
        tk.Label(self._desc_frame, text="Description :", bg=bg_color,
                 font=("Arial", 9)).pack(anchor="w")
        self.desc_text = tk.Text(self._desc_frame, height=4, width=72,
                                  font=("Arial", 9), relief=tk.SOLID, bd=1)
        self.desc_text.pack(fill=tk.X)

    def _toggle_desc(self):
        if self._desc_visible:
            self._desc_frame.pack_forget()
            self._desc_btn.configure(text="+ description")
        else:
            self._desc_frame.pack(fill=tk.X, padx=30, pady=(2, 4))
            self._desc_btn.configure(text="– description")
        self._desc_visible = not self._desc_visible

    def get_code(self):
        v = self.project_var.get()
        return v.split("  –  ")[0].strip() if v else None

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


class App(tk.Tk):
    ROW_COLORS = ("#ffffff", "#f5f7ff")

    def __init__(self):
        super().__init__()
        self.title("Générateur de Factures – Bpifrance")
        self.configure(bg="#f2f4f8")
        self.resizable(True, True)
        self.minsize(820, 420)

        self._rows = []
        self._output_dir = (DEFAULT_DIR if os.path.isdir(DEFAULT_DIR)
                            else os.path.expanduser("~"))
        self._build()

    # ── Construction UI ────────────────────────────────────────────────────

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg="#003189", pady=14)
        hdr.pack(fill=tk.X)
        tk.Label(hdr,
                 text="Générateur de Factures  ·  Pierre-Alexis Fillon / Bpifrance",
                 font=("Arial", 14, "bold"), fg="white", bg="#003189").pack()
        tk.Label(hdr, text="Prix\u00a0HT\u00a0: 240\u00a0€ / jour  ·  TVA non applicable",
                 font=("Arial", 10), fg="#aabfff", bg="#003189").pack()

        # Settings bar
        bar = tk.Frame(self, bg="#dde3f0", pady=8, padx=14)
        bar.pack(fill=tk.X)

        today = date.today()

        tk.Label(bar, text="Mois :", bg="#dde3f0",
                 font=("Arial", 10)).pack(side=tk.LEFT)
        month_names = [f"{i:02d} – {MONTHS_FR[i].capitalize()}"
                       for i in range(1, 13)]
        self._month_cb = ttk.Combobox(bar, values=month_names,
                                       state="readonly", width=15,
                                       font=("Arial", 10))
        self._month_cb.current(today.month - 1)
        self._month_cb.pack(side=tk.LEFT, padx=(4, 12))

        tk.Label(bar, text="Année :", bg="#dde3f0",
                 font=("Arial", 10)).pack(side=tk.LEFT)
        self._year_var = tk.StringVar(value=str(today.year))
        ttk.Spinbox(bar, from_=2020, to=2035, textvariable=self._year_var,
                    width=7, font=("Arial", 10)).pack(side=tk.LEFT, padx=(4, 12))

        tk.Label(bar, text="Jour de facturation :", bg="#dde3f0",
                 font=("Arial", 10)).pack(side=tk.LEFT)
        self._day_var = tk.StringVar(value=str(today.day))
        ttk.Spinbox(bar, from_=1, to=31, textvariable=self._day_var,
                    width=5, font=("Arial", 10)).pack(side=tk.LEFT, padx=(4, 16))

        tk.Label(bar, text="Sortie :", bg="#dde3f0",
                 font=("Arial", 10)).pack(side=tk.LEFT)
        self._dir_lbl = tk.Label(bar, text=self._short(self._output_dir),
                                  bg="#dde3f0", fg="#333", font=("Arial", 9))
        self._dir_lbl.pack(side=tk.LEFT, padx=4)
        tk.Button(bar, text="Choisir…", command=self._pick_dir,
                  bg="#c0cde8", relief=tk.FLAT, padx=6).pack(side=tk.LEFT)

        # Rows area
        lf = tk.LabelFrame(self, text="  Factures à générer  ",
                            bg="#f2f4f8", font=("Arial", 10, "bold"),
                            padx=10, pady=8)
        lf.pack(fill=tk.BOTH, expand=True, padx=14, pady=10)

        canvas = tk.Canvas(lf, bg="#f2f4f8", highlightthickness=0)
        sb = ttk.Scrollbar(lf, orient="vertical", command=canvas.yview)
        self._rows_frame = tk.Frame(canvas, bg="#f2f4f8")
        self._rows_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom buttons
        btns = tk.Frame(self, bg="#f2f4f8", pady=8, padx=14)
        btns.pack(fill=tk.X)
        tk.Button(btns, text="+ Ajouter une facture", command=self._add_row,
                  bg="#4CAF50", fg="white", relief=tk.FLAT,
                  font=("Arial", 10, "bold"), padx=12, pady=6).pack(side=tk.LEFT)
        tk.Button(btns, text="  Générer les PDF  ", command=self._generate,
                  bg="#003189", fg="white", relief=tk.FLAT,
                  font=("Arial", 12, "bold"), padx=20, pady=6).pack(side=tk.RIGHT)

        self._add_row()   # row par défaut

    # ── Helpers ────────────────────────────────────────────────────────────

    def _short(self, path, n=50):
        return path if len(path) <= n else "…" + path[-(n - 1):]

    def _pick_dir(self):
        d = filedialog.askdirectory(initialdir=self._output_dir)
        if d:
            self._output_dir = d
            self._dir_lbl.configure(text=self._short(d))

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
                    w.configure(text=f"#{i}")
                    break

    def _get_month(self):
        return int(self._month_cb.get().split(" – ")[0])

    # ── Génération ─────────────────────────────────────────────────────────

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

        generated, errors = [], []

        for row in valid:
            code  = row.get_code()
            name  = row.get_name()
            days  = row.get_days()
            xdesc = row.get_extra_desc()
            fname = f"Pierre-Alexis Fillon - {code} - {month:02d} - {year}.pdf"
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
                    extra_desc=xdesc,
                )
                generated.append(fname)
            except Exception as exc:
                errors.append(f"{code}: {exc}")

        if generated:
            lines = [f"✓  {f}" for f in generated]
            if errors:
                lines += ["", "Erreurs :"] + [f"✗  {e}" for e in errors]
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
    # Check Python version
    if sys.version_info < (3, 8):
        print("Python 3.8+ requis.")
        sys.exit(1)

    app = App()
    app.mainloop()
