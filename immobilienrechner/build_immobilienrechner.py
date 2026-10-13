#!/usr/bin/env python3
"""
Generator für den Immobilien-Rechner Deutschland (Stand 2026).

Erzeugt eine reine .xlsx-Datei (keine Makros) mit Formeln, benannten Bereichen,
Datenüberprüfung und Blattschutz (ohne Passwort).

Aufbau:
  * Jede Eingabe, jeder Parameter und jede Zwischengröße ist ein benannter Bereich.
  * Formeln werden als Vorlagen geschrieben:
        {Name}     -> skalare Größe (benannter Bereich)
        [key]      -> Spalte des Jahresmodells, gleiche Zeile
        [key@-1]   -> Spalte des Jahresmodells, Vorjahr
        [[key]]    -> komplette Spalte des Jahresmodells (Jahr 0..40)
  * Das Jahresmodell (Blatt "Cashflow-Plan") wird für Szenarien und
    Sensitivitäten auf dem ausgeblendeten Blatt "_Sens" erneut erzeugt
    (Schattenrechnung), statt Excel-Datentabellen zu verwenden. So rechnet
    die Datei in Excel 365, Excel Online und LibreOffice identisch.

Aufruf:  python3 build_immobilienrechner.py [ausgabe.xlsx] [Name=Wert ...]
"""
import datetime as dt
import re
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

T_MAX = 40  # maximale Anzahl Modelljahre

# ---------------------------------------------------------------- Listen-Texte
KA = "Kapitalanlage"
EN = "Eigennutzung"
JA, NEIN = "Ja", "Nein"
TARIF = "Aus zvE berechnen (§32a)"
DIREKT = "Grenzsteuersatz direkt"
ETW = "Eigentumswohnung"
HAUS = "Haus (EFH/DHH/RH)"
MFH = "Mehrfamilienhaus"
AUF_AH = "Arbeitshilfe (berechnet)"
AUF_MAN = "Manuell (Kaufvertrag/Gutachten)"
ERH = "Erhaltungsaufwand"
HK = "Herstellungsaufwand"
INST_P = "Vorschlag (Peters'sche Formel)"
INST_M = "Manuell"
NB_FIX = "Fester Betrag"
NB_ZIEL = "Bis Ziel-Beleihungsauslauf"
NB_MOD = "Modernisierung dieses Objekts"
NB_FOLGE = "Eigenkapital für weiteres Mietobjekt"
NB_ETF = "ETF/Wertpapiere"
NB_KONSUM = "Privat/Konsum"
ZUS = "Zusammenveranlagung"
EINZ = "Einzelveranlagung"
DENK_NEIN = "Nein"
DENK_7I = "Baudenkmal (§7i / §10f)"
DENK_7H = "Sanierungsgebiet (§7h / §10f)"
ZUSTAENDE = ["Neubau / neuwertig", "Modernisiert", "Gepflegt (Standard)",
             "Renovierungsbedürftig", "Sanierungsbedürftig"]
LAENDER = [  # Land, GrESt, KiSt
    ("Baden-Württemberg", 0.05, 0.08), ("Bayern", 0.035, 0.08), ("Berlin", 0.06, 0.09),
    ("Brandenburg", 0.065, 0.09), ("Bremen", 0.055, 0.09), ("Hamburg", 0.055, 0.09),
    ("Hessen", 0.06, 0.09), ("Mecklenburg-Vorpommern", 0.06, 0.09),
    ("Niedersachsen", 0.05, 0.09), ("Nordrhein-Westfalen", 0.065, 0.09),
    ("Rheinland-Pfalz", 0.05, 0.09), ("Saarland", 0.065, 0.09), ("Sachsen", 0.055, 0.09),
    ("Sachsen-Anhalt", 0.05, 0.09), ("Schleswig-Holstein", 0.065, 0.09),
    ("Thüringen", 0.05, 0.09)]
KENNZAHLEN = ["IRR nach Steuern", "Vermögensvorteil ggü. Alternative",
              "Cashflow n. St. Jahr 1 (mtl.)", "Vermögen am Ende (Wert − Schuld + kum. CF)"]

# ---------------------------------------------------------------- Formate/Stile
FONT = "Arial"
F_EUR = '#,##0 "€";[Red]-#,##0 "€";"-"'
F_EUR2 = '#,##0.00 "€";[Red]-#,##0.00 "€";"-"'
F_PCT = '0.0%;[Red]-0.0%;"-"'
F_PCT2 = '0.00%;[Red]-0.00%;"0,00 %"'
F_PCTI = '0.00%'
F_INT = '0'
F_NUM = '#,##0'
F_NUM1 = '#,##0.0'
F_NUM2 = '#,##0.00'
F_M2 = '#,##0.0 "m²"'
F_FAKT = '0.0"x"'
F_DATE = 'DD.MM.YYYY'
F_TXT = '@'

C_IN = "FFF2CC"      # Eingabe (gelb)
C_HEAD = "1F4E78"    # Kopfzeile
C_SUB = "D9E1F2"     # Abschnitt
C_KEY = "DDEBF7"     # Kernergebnis
C_ORANGE = "F8CBAD"  # unsicherer Parameter
C_GREY = "F2F2F2"
C_GREEN = "C6EFCE"
C_YELLOW = "FFEB9C"
C_RED = "FFC7CE"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fill(c):
    return PatternFill("solid", fgColor=c)


def fnt(**kw):
    kw.setdefault("size", 10)
    return Font(name=FONT, **kw)


NAME_BAD = re.compile(r"^([A-Za-z]{1,3}\d|[RrCc]\d*$|[RrCc]\d+[CcRr])")


class Model:
    def __init__(self, overrides=None):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.nodes = {}       # name -> {'tpl':..., 'kind':...}
        self.ov = overrides or {}
        self.dv = {}          # (sheet, key) -> DataValidation
        self.engine = []      # Spaltendefinitionen des Jahresmodells
        self.eng_col = {}     # key -> Spaltenbuchstabe auf Cashflow-Plan
        self.eng_row0 = None  # Zeile von Jahr 0
        self.sheets = {}
        self.used = set()     # in Formeln verwendete Namen (Prüfung am Ende)

    # ------------------------------------------------------------ Hilfen
    def sheet(self, key, title):
        ws = self.wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 90
        self.sheets[key] = ws
        return ws

    def define(self, name, ws, ref):
        assert not NAME_BAD.match(name), f"Ungültiger Name {name}"
        low = name.lower()
        for n in self.wb.defined_names:
            assert n.lower() != low, f"Name doppelt: {name}"
        self.wb.defined_names[name] = DefinedName(name, attr_text=f"'{ws.title}'!{ref}")

    def leaf(self, name, kind="table"):
        self.nodes[name] = {"tpl": None, "kind": kind}

    # Vorlage -> Excel-Formel (Hauptmodell)
    def to_main(self, tpl, row=None):
        s = re.sub(r"\[\[(\w+)\]\]", lambda m: "Z_" + m.group(1), tpl)
        if row is not None:
            s = re.sub(r"\[(\w+)@-1\]", lambda m: f"{self.eng_col[m.group(1)]}{row - 1}", s)
            s = re.sub(r"\[(\w+)\]", lambda m: f"{self.eng_col[m.group(1)]}{row}", s)
        s = re.sub(r"\{(\w+)\}", lambda m: self._chk(m.group(1)), s)
        assert "[" not in s and "{" not in s, s
        return "=" + s

    def _chk(self, n):
        self.used.add(n)
        return n

    def style(self, c, fmt=None, bold=False, fillc=None, color=None, wrap=False,
              align=None, size=10, italic=False, border=True):
        c.font = fnt(bold=bold, color=color, size=size, italic=italic)
        if fmt:
            c.number_format = fmt
        if fillc:
            c.fill = fill(fillc)
        if border:
            c.border = BORDER
        c.alignment = Alignment(wrap_text=wrap, vertical="center",
                                horizontal=align)

    def add_dv(self, ws, cell, spec, hint):
        kind = spec[0]
        key = (ws.title,) + tuple(spec) + (hint or "",)
        if key not in self.dv:
            if kind == "list":
                dv = DataValidation(type="list", formula1=spec[1], allow_blank=False)
                dv.error = "Bitte einen Eintrag aus der Auswahlliste wählen."
            elif kind in ("dec", "int"):
                dv = DataValidation(type="decimal" if kind == "dec" else "whole",
                                    operator="between", formula1=str(spec[1]),
                                    formula2=str(spec[2]), allow_blank=False)
                dv.error = spec[3]
            elif kind == "date":
                dv = DataValidation(type="date", operator="between",
                                    formula1="1", formula2="73415", allow_blank=False)
                dv.error = "Bitte ein gültiges Datum (TT.MM.JJJJ) eingeben."
            elif kind == "text":
                dv = DataValidation(type="textLength", operator="lessThanOrEqual",
                                    formula1="60", allow_blank=True)
                dv.error = "Maximal 60 Zeichen."
            if hint:
                dv.promptTitle = "Hinweis"
                dv.prompt = hint if len(hint) <= 250 else hint[:247] + "…"
                dv.showInputMessage = True
            dv.errorTitle = "Ungültige Eingabe"
            dv.showErrorMessage = True
            dv.errorStyle = "stop"
            ws.add_data_validation(dv)
            self.dv[key] = dv
        self.dv[key].add(cell.coordinate)

    # ------------------------------------------------------------ Zellen
    def inp(self, ws, row, name, label, default, fmt, spec, hint, col=3, unit=None,
            write_label=True):
        val = self.ov.get(name, default)
        c = ws.cell(row=row, column=col, value=val)
        self.style(c, fmt=fmt, fillc=C_IN, color="0000FF")
        c.protection = Protection(locked=False)
        if write_label:
            self.label(ws, row, label, unit, hint)
        self.add_dv(ws, c, spec, hint)
        self.define(name, ws, f"${get_column_letter(col)}${row}")
        self.nodes[name] = {"tpl": None, "kind": "input"}
        return c

    def label(self, ws, row, label, unit=None, hint=None, hint_col=7):
        a = ws.cell(row=row, column=2, value=label)
        self.style(a, wrap=True)
        if unit is not None:
            u = ws.cell(row=row, column=4, value=unit)
            self.style(u, color="7F7F7F")
        if hint:
            h = ws.cell(row=row, column=hint_col, value=hint)
            self.style(h, wrap=True, color="404040", size=9)

    def calc(self, ws, row, name, label, tpl, fmt, hint=None, col=3, key=False,
             unit=None, write_label=True):
        self.nodes[name] = {"tpl": tpl, "kind": "calc"}
        c = ws.cell(row=row, column=col, value=self.to_main(tpl))
        self.style(c, fmt=fmt, fillc=C_KEY if key else None, bold=key)
        if write_label and label is not None:
            self.label(ws, row, label, unit, hint)
        if name:
            self.define(name, ws, f"${get_column_letter(col)}${row}")
        return c

    def view(self, ws, row, col, formula, fmt=None, bold=False, fillc=None, color=None,
             wrap=False, align=None):
        """Anzeigezelle ohne eigenen Namen (Formel aus Vorlage)."""
        c = ws.cell(row=row, column=col,
                    value=self.to_main(formula) if isinstance(formula, str) and not formula.startswith("'") else formula)
        self.style(c, fmt=fmt, bold=bold, fillc=fillc, color=color, wrap=wrap, align=align)
        return c

    def text(self, ws, row, col, value, bold=False, size=10, color=None, fillc=None,
             wrap=False, border=False, italic=False, align=None):
        c = ws.cell(row=row, column=col, value=value)
        self.style(c, bold=bold, size=size, color=color, fillc=fillc, wrap=wrap,
                   border=border, italic=italic, align=align)
        return c

    def section(self, ws, row, title, ncols=7):
        for col in range(2, ncols + 1):
            c = ws.cell(row=row, column=col)
            c.fill = fill(C_HEAD)
            c.border = BORDER
        c = ws.cell(row=row, column=2, value=title)
        c.font = fnt(bold=True, color="FFFFFF", size=11)
        c.fill = fill(C_HEAD)
        ws.row_dimensions[row].height = 20

    def subhead(self, ws, row, values, start=2, fillc=C_SUB):
        for i, v in enumerate(values):
            c = ws.cell(row=row, column=start + i, value=v)
            self.style(c, bold=True, fillc=fillc, wrap=True, align="center")

    def title(self, ws, text, sub=None):
        ws["B1"] = text
        ws["B1"].font = fnt(bold=True, size=16, color=C_HEAD)
        if sub:
            ws["B2"] = sub
            ws["B2"].font = fnt(italic=True, size=9, color="595959")
        ws.column_dimensions["A"].width = 2

    def protect(self, ws):
        ws.protection.sheet = True
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False


# ====================================================================== Tarif
def T_tpl(x):
    """Tarifliche ESt nach §32a EStG (inkl. Splitting) als Vorlage."""
    v = f"INT(({x})/{{Split}})"
    inner = (f"IF({v}<={{P_GFB}},0,IF({v}<={{P_Z2}},"
             f"({{P_Tarif_a1}}*({v}-{{P_GFB}})/10000+{{P_Tarif_b1}})*({v}-{{P_GFB}})/10000,"
             f"IF({v}<={{P_Z3}},({{P_Tarif_a2}}*({v}-{{P_Z2}})/10000+{{P_Tarif_b2}})*({v}-{{P_Z2}})/10000+{{P_Tarif_c2}},"
             f"IF({v}<={{P_Z4}},{{P_Tarif_s4}}*{v}-{{P_Tarif_c4}},{{P_Tarif_s5}}*{v}-{{P_Tarif_c5}}))))")
    return f"{{Split}}*INT({inner})"


def soli_tpl(e):
    return f"IF({e}<={{Soli_FG}},0,MIN({{P_Soli_Satz}}*{e},{{P_Soli_Milderung}}*({e}-{{Soli_FG}})))"


def total_tax_tpl(e):
    return f"{e}+{soli_tpl(e)}+{e}*{{KiSt_Satz_eff}}"


# ====================================================================== Blätter
def build_parameter(m):
    ws = m.sheet("par", "Parameter")
    m.title(ws, "Parameter (gesetzliche Werte & Annahmen)",
            "Alle gesetzlichen Werte zentral. Gelbe Zellen können bei Gesetzesänderungen angepasst werden. "
            "Spalte 'Status': PRÜFEN = Wert nicht sicher belegt bzw. Annahme – bitte selbst verifizieren.")
    for col, w in zip("BCDEFG", (46, 14, 10, 62, 12, 12)):
        ws.column_dimensions[col].width = w
    r = 4
    status_cells = []

    def p(name, label, val, fmt, unit, quelle, stand="10/2026", sicher=True, spec=None):
        nonlocal r
        spec = spec or ("dec", -1e9, 1e9, "Bitte eine Zahl eingeben.")
        m.inp(ws, r, name, label, val, fmt, spec, None, unit=unit)
        q = m.text(ws, r, 5, quelle, wrap=True, border=True, size=9)
        m.text(ws, r, 6, stand, border=True, size=9)
        s = m.text(ws, r, 7, "sicher" if sicher else "PRÜFEN", border=True, bold=not sicher,
                   fillc=None if sicher else C_ORANGE, size=9, align="center")
        r += 1

    def sec(t):
        nonlocal r
        r += 1
        m.section(ws, r, t)
        m.subhead(ws, r + 1, ["Parameter", "Wert", "Einheit", "Quelle / Rechtsgrundlage", "Stand", "Status"])
        r += 2

    sec("Einkommensteuertarif 2026 (§32a EStG)")
    p("P_GFB", "Grundfreibetrag (Ende Zone 1)", 12348, F_NUM, "€", "§32a Abs. 1 Nr. 1 EStG i.d.F. Steuerfortentwicklungsgesetz (BGBl. 2024 I Nr. 449)")
    p("P_Z2", "Ende Zone 2", 17799, F_NUM, "€", "§32a Abs. 1 Nr. 2 EStG")
    p("P_Z3", "Ende Zone 3", 69878, F_NUM, "€", "§32a Abs. 1 Nr. 3 EStG")
    p("P_Z4", "Ende Zone 4 (Beginn Reichensteuer)", 277825, F_NUM, "€", "§32a Abs. 1 Nr. 4 EStG")
    p("P_Tarif_a1", "Zone 2: Faktor y²", 914.51, F_NUM2, "", "§32a Abs. 1 Nr. 2 EStG: (914,51·y + 1.400)·y")
    p("P_Tarif_b1", "Zone 2: Faktor y", 1400, F_NUM2, "", "§32a Abs. 1 Nr. 2 EStG")
    p("P_Tarif_a2", "Zone 3: Faktor z²", 173.10, F_NUM2, "", "§32a Abs. 1 Nr. 3 EStG: (173,10·z + 2.397)·z + 1.034,87")
    p("P_Tarif_b2", "Zone 3: Faktor z", 2397, F_NUM2, "", "§32a Abs. 1 Nr. 3 EStG")
    p("P_Tarif_c2", "Zone 3: Konstante", 1034.87, F_NUM2, "€", "§32a Abs. 1 Nr. 3 EStG")
    p("P_Tarif_s4", "Zone 4: Steuersatz", 0.42, F_PCT, "", "§32a Abs. 1 Nr. 4 EStG: 0,42·x − 11.135,63")
    p("P_Tarif_c4", "Zone 4: Abzug", 11135.63, F_NUM2, "€", "§32a Abs. 1 Nr. 4 EStG")
    p("P_Tarif_s5", "Zone 5: Steuersatz", 0.45, F_PCT, "", "§32a Abs. 1 Nr. 5 EStG: 0,45·x − 19.470,38")
    p("P_Tarif_c5", "Zone 5: Abzug", 19470.38, F_NUM2, "€", "§32a Abs. 1 Nr. 5 EStG")
    p("P_Soli_Satz", "Solidaritätszuschlag", 0.055, F_PCT, "", "§4 SolZG")
    p("P_Soli_Freigrenze", "Soli-Freigrenze (Einzelveranlagung; Splitting ×2)", 20350, F_NUM, "€ ESt", "§3 Abs. 3 SolZG i.d.F. Steuerfortentwicklungsgesetz")
    p("P_Soli_Milderung", "Soli-Milderungszone", 0.119, F_PCT, "", "§4 S. 2 SolZG")

    sec("Abschreibung (AfA) & Sonderabschreibungen")
    p("P_AfA_vor1925", "Lineare AfA Fertigstellung vor 1925", 0.025, F_PCT, "p.a.", "§7 Abs. 4 S. 1 Nr. 2 b EStG")
    p("P_AfA_1925_2022", "Lineare AfA Fertigstellung 1925–2022", 0.02, F_PCT, "p.a.", "§7 Abs. 4 S. 1 Nr. 2 a EStG")
    p("P_AfA_ab2023", "Lineare AfA Fertigstellung ab 2023", 0.03, F_PCT, "p.a.", "§7 Abs. 4 S. 1 Nr. 2 a EStG (JStG 2022)")
    p("P_AfA_Grenze1", "Grenzjahr 2,5 % → 2 %", 1925, F_INT, "Jahr", "§7 Abs. 4 EStG")
    p("P_AfA_Grenze2", "Grenzjahr 2 % → 3 %", 2023, F_INT, "Jahr", "§7 Abs. 4 EStG")
    p("P_Degr_Satz", "Degressive AfA Wohngebäude", 0.05, F_PCT, "v. Restwert", "§7 Abs. 5a EStG (Wachstumschancengesetz)")
    p("P_Degr_Beginn", "Degressive AfA: Baubeginn ab", dt.date(2023, 10, 1), F_DATE, "Datum", "§7 Abs. 5a EStG", spec=("date",))
    p("P_Degr_Ende", "Degressive AfA: Baubeginn bis", dt.date(2029, 9, 30), F_DATE, "Datum", "§7 Abs. 5a EStG; Erwerb nur bis Ende des Fertigstellungsjahres", sicher=False, spec=("date",))
    p("P_7b_Satz", "Sonder-AfA §7b je Jahr", 0.05, F_PCT, "p.a.", "§7b Abs. 1 EStG")
    p("P_7b_Jahre", "Sonder-AfA §7b Dauer", 4, F_INT, "Jahre", "§7b Abs. 1 EStG (Jahr der Anschaffung + 3 Folgejahre)")
    p("P_7b_Antrag_Beginn", "§7b: Bauantrag ab", dt.date(2023, 1, 1), F_DATE, "Datum", "§7b Abs. 2 Nr. 1 EStG", spec=("date",))
    p("P_7b_Antrag_Ende", "§7b: Bauantrag bis", dt.date(2029, 9, 30), F_DATE, "Datum", "§7b Abs. 2 EStG i.d.F. Wachstumschancengesetz – Quellen teils widersprüchlich", sicher=False, spec=("date",))
    p("P_7b_BKO", "§7b: Baukostenobergrenze", 5200, F_NUM, "€/m²", "§7b Abs. 2 Nr. 2 EStG (Bauantrag ab 2023) – bitte verifizieren", sicher=False)
    p("P_7b_BMG_max", "§7b: max. Bemessungsgrundlage", 4000, F_NUM, "€/m²", "§7b Abs. 3 Nr. 2 EStG (Bauantrag ab 2023)")
    p("P_7i_Satz1", "§7i/§7h: Satz Jahre 1–8", 0.09, F_PCT, "p.a.", "§7i, §7h EStG")
    p("P_7i_Jahre1", "§7i/§7h: Jahre Phase 1", 8, F_INT, "Jahre", "§7i, §7h EStG")
    p("P_7i_Satz2", "§7i/§7h: Satz Jahre 9–12", 0.07, F_PCT, "p.a.", "§7i, §7h EStG")
    p("P_7i_Jahre2", "§7i/§7h: Jahre Phase 2", 4, F_INT, "Jahre", "§7i, §7h EStG")
    p("P_10f_Satz", "§10f (Eigennutzung Denkmal/Sanierung)", 0.09, F_PCT, "p.a.", "§10f EStG")
    p("P_10f_Jahre", "§10f Dauer", 10, F_INT, "Jahre", "§10f EStG")
    p("P_AN_Quote", "Anschaffungsnahe HK: Grenze", 0.15, F_PCT, "v. Geb.-AK", "§6 Abs. 1 Nr. 1a EStG (netto, ohne USt)")
    p("P_AN_Jahre", "Anschaffungsnahe HK: Zeitraum", 3, F_INT, "Jahre", "§6 Abs. 1 Nr. 1a EStG")
    p("P_USt", "Umsatzsteuer (für 15-%-Prüfung netto)", 0.19, F_PCT, "", "§12 UStG")
    p("P_82b_max", "§82b EStDV: max. Verteilungsjahre", 5, F_INT, "Jahre", "§82b EStDV (2–5 Jahre)")
    p("P_Inventar_ND_Hinweis", "Richtwert Nutzungsdauer Einbauküche (nur Info, Eingabe auf 'Eingaben')", 10, F_INT, "Jahre", "BFH 3.8.2016, IX R 14/15")
    p("P_Disagio_max", "Disagio sofort abziehbar bis", 0.05, F_PCT, "", "BMF 20.10.2003 (marktüblich, Zinsbindung ≥ 5 J.)")
    p("P_Disagio_ZB_min", "… bei Zinsbindung mind.", 5, F_INT, "Jahre", "BMF 20.10.2003")
    p("P_km_Satz", "Fahrten zum Mietobjekt", 0.30, F_EUR2, "€/km", "§9 Abs. 1 S. 3 Nr. 4a EStG analog (Reisekosten)", sicher=False)
    p("P_Konto_Pausch", "Kontoführung (Nichtbeanstandung)", 16, F_EUR, "€/Jahr", "Verwaltungspraxis (ohne Nachweis)", sicher=False)

    sec("Vermietung: Grenzen & Prüfungen")
    p("P_Verb_Voll", "Verbilligte Vermietung: voll entgeltlich ab", 0.66, F_PCT, "d. ortsübl. Miete", "§21 Abs. 2 S. 1 EStG")
    p("P_Verb_Min", "Verbilligte Vermietung: Aufteilung unter", 0.50, F_PCT, "d. ortsübl. Miete", "§21 Abs. 2 S. 2 EStG (50–66 %: Totalüberschussprognose)")
    p("P_TUEP_Jahre", "Totalüberschussprognose: Zeitraum", 30, F_INT, "Jahre", "BMF 8.10.2004, BStBl. I 2004, 933")
    p("P_Spekfrist", "Spekulationsfrist §23", 10, F_INT, "Jahre", "§23 Abs. 1 Nr. 1 EStG")
    p("P_23_Freigrenze", "Freigrenze private Veräußerungsgeschäfte", 1000, F_NUM, "€/Jahr", "§23 Abs. 3 S. 5 EStG (ab VZ 2024)")
    p("P_DreiObj", "Drei-Objekt-Grenze: max. Objekte", 3, F_INT, "Objekte", "BMF 26.3.2004 (gewerbl. Grundstückshandel)")
    p("P_DreiObj_Jahre", "Drei-Objekt-Grenze: Zeitraum", 5, F_INT, "Jahre", "BMF 26.3.2004")
    p("P_Kuend_489", "Sonderkündigungsrecht Darlehen nach", 10, F_INT, "Jahre", "§489 Abs. 1 Nr. 2 BGB (keine Vorfälligkeitsentschädigung)")

    sec("Eigennutzung: Steuerermäßigungen")
    p("P_35a_HW_Satz", "§35a Handwerkerleistungen: Satz", 0.20, F_PCT, "d. Lohnkosten", "§35a Abs. 3 EStG")
    p("P_35a_HW_Max", "§35a Handwerker: max. Ermäßigung", 1200, F_NUM, "€/Jahr", "§35a Abs. 3 EStG")
    p("P_35a_HHDL_Satz", "§35a haushaltsnahe Dienstl.: Satz", 0.20, F_PCT, "d. Lohnkosten", "§35a Abs. 2 EStG")
    p("P_35a_HHDL_Max", "§35a haushaltsnahe Dienstl.: max.", 4000, F_NUM, "€/Jahr", "§35a Abs. 2 EStG")
    p("P_35c_S1", "§35c energ. Sanierung: Jahr 1", 0.07, F_PCT, "", "§35c Abs. 1 EStG")
    p("P_35c_S2", "§35c: Jahr 2", 0.07, F_PCT, "", "§35c Abs. 1 EStG")
    p("P_35c_S3", "§35c: Jahr 3", 0.06, F_PCT, "", "§35c Abs. 1 EStG")
    p("P_35c_MaxAufw", "§35c: max. begünstigte Aufwendungen", 200000, F_NUM, "€/Objekt", "§35c Abs. 1 EStG (max. 40.000 € Ermäßigung)")
    p("P_35c_Alter", "§35c: Mindestalter Gebäude", 10, F_INT, "Jahre", "§35c Abs. 1 S. 2 EStG")
    p("P_35c_Ende", "§35c: Abschluss der Maßnahme bis Jahr", 2029, F_INT, "Jahr", "§35c Abs. 1 / §52 Abs. 35a EStG (Beginn nach 2019, Abschluss vor 2030)")

    sec("Kapitalanlage-Alternative (ETF)")
    p("P_AbgSt", "Abgeltungsteuer", 0.25, F_PCT, "", "§32d Abs. 1 EStG")
    p("P_TFS", "Teilfreistellung Aktienfonds", 0.30, F_PCT, "", "§20 Abs. 1 InvStG")
    p("P_SPB", "Sparer-Pauschbetrag je Person", 1000, F_NUM, "€", "§20 Abs. 9 EStG")
    p("P_Basiszins", "Basiszins Vorabpauschale 2026 (nur Info)", 0.032, F_PCTI, "", "BMF 13.1.2026 (§18 Abs. 4 InvStG)")

    sec("Kaufpreisaufteilung (vereinfachte Logik der BMF-Arbeitshilfe, Sachwertverfahren)")
    p("P_BPI", "Baupreisindex Wohngebäude (2010 = 100)", 190, F_NUM1, "", "BMF 20.1.2025: 183,3–186,7 für 2025; 2026 geschätzt (Destatis +~3 %)", sicher=False)
    p("P_GND", "Gesamtnutzungsdauer Wohngebäude", 80, F_INT, "Jahre", "Anlage 1 ImmoWertV 2021")
    p("P_RND_min", "Mindest-Restnutzungsdauer (Anteil GND)", 0.25, F_PCT, "", "Annahme (vereinfachte Modernisierungslogik)", sicher=False)
    p("P_Peters_Faktor", "Peters'sche Formel: Faktor", 1.5, F_NUM2, "", "Peters'sche Formel (HK × 1,5 / 80 Jahre)")
    p("P_Peters_Jahre", "Peters'sche Formel: Nutzungsjahre", 80, F_INT, "Jahre", "Peters'sche Formel")
    p("P_Peters_ETW", "ETW: Anteil Sondereigentum (eigene Reserve)", 0.30, F_PCT, "", "Faustregel 65–70 % über WEG-Rücklage, 30–35 % Sondereigentum", sicher=False)
    r += 1
    m.subhead(ws, r, ["Objekttyp", "NHK 2010 Stufe 3 (€/m² BGF)", "BGF je m² Wohnfl.", "Quelle", "", "Status"])
    r += 1
    nhk = [(ETW, 825, 1.40), (HAUS, 835, 1.70), (MFH, 825, 1.40)]
    r0 = r
    for t, v, f in nhk:
        m.text(ws, r, 2, t, border=True)
        for col, val, fmtx in ((3, v, F_NUM), (4, f, F_NUM2)):
            c = ws.cell(row=r, column=col, value=val)
            m.style(c, fmt=fmtx, fillc=C_IN, color="0000FF")
            c.protection = Protection(locked=False)
            m.add_dv(ws, c, ("dec", 0, 100000, "Bitte eine positive Zahl eingeben."), None)
        m.text(ws, r, 5, "NHK 2010 (Anlage 4 ImmoWertV 2021): Typ 1.01 bzw. 4.1; BGF-Faktor = eigene Annahme", border=True, size=9, wrap=True)
        m.text(ws, r, 7, "PRÜFEN", border=True, bold=True, fillc=C_ORANGE, size=9, align="center")
        r += 1
    m.define("Tab_Objekttyp", ws, f"$B${r0}:$B${r-1}")
    m.define("Tab_NHK", ws, f"$C${r0}:$C${r-1}")
    m.define("Tab_BGF", ws, f"$D${r0}:$D${r-1}")
    for n in ("Tab_Objekttyp", "Tab_NHK", "Tab_BGF"):
        m.leaf(n)
    r += 1
    m.subhead(ws, r, ["Ausstattungsstandard (NHK-Stufe)", "Faktor ggü. Stufe 3", "", "Quelle", "", "Status"])
    r += 1
    r0 = r
    for st, f in ((1, 0.78), (2, 0.87), (3, 1.0), (4, 1.20), (5, 1.51)):
        m.text(ws, r, 2, f"Stufe {st}", border=True)
        c = ws.cell(row=r, column=3, value=f)
        m.style(c, fmt=F_NUM2, fillc=C_IN, color="0000FF")
        c.protection = Protection(locked=False)
        m.add_dv(ws, c, ("dec", 0, 10, "Bitte einen Faktor zwischen 0 und 10 eingeben."), None)
        m.text(ws, r, 5, "Verhältnis der NHK-2010-Werte Typ 1.01 (655/725/835/1.005/1.260 €/m²)", border=True, size=9)
        m.text(ws, r, 7, "PRÜFEN", border=True, bold=True, fillc=C_ORANGE, size=9, align="center")
        r += 1
    m.define("Tab_Standard", ws, f"$C${r0}:$C${r-1}")
    m.leaf("Tab_Standard")
    r += 1
    m.subhead(ws, r, ["Zustand", "Verjüngung RND (Jahre)", "Faktor Instandhaltung", "Erläuterung", "", "Status"])
    r += 1
    r0 = r
    for z, verj, fi in zip(ZUSTAENDE, (0, 15, 0, -5, -10), (0.8, 0.9, 1.0, 1.25, 1.5)):
        m.text(ws, r, 2, z, border=True)
        for col, val, fmtx in ((3, verj, F_INT), (4, fi, F_NUM2)):
            c = ws.cell(row=r, column=col, value=val)
            m.style(c, fmt=fmtx, fillc=C_IN, color="0000FF")
            c.protection = Protection(locked=False)
            m.add_dv(ws, c, ("dec", -80, 80, "Bitte eine Zahl zwischen −80 und 80 eingeben."), None)
        m.text(ws, r, 5, "Eigene Annahme: verlängert/verkürzt die Restnutzungsdauer und skaliert die Instandhaltungsreserve", border=True, size=9, wrap=True)
        m.text(ws, r, 7, "PRÜFEN", border=True, bold=True, fillc=C_ORANGE, size=9, align="center")
        r += 1
    m.define("Liste_Zustand", ws, f"$B${r0}:$B${r-1}")
    m.define("Tab_Zust_Verj", ws, f"$C${r0}:$C${r-1}")
    m.define("Tab_Zust_Inst", ws, f"$D${r0}:$D${r-1}")
    for n in ("Liste_Zustand", "Tab_Zust_Verj", "Tab_Zust_Inst"):
        m.leaf(n)

    r += 1
    m.section(ws, r, "Grunderwerbsteuer & Kirchensteuer je Bundesland")
    r += 1
    m.subhead(ws, r, ["Bundesland", "Grunderwerbsteuer", "Kirchensteuer", "Quelle", "Stand", "Status"])
    r += 1
    r0 = r
    for land, g, k in LAENDER:
        m.text(ws, r, 2, land, border=True)
        for col, val in ((3, g), (4, k)):
            c = ws.cell(row=r, column=col, value=val)
            m.style(c, fmt=F_PCT, fillc=C_IN, color="0000FF")
            c.protection = Protection(locked=False)
            m.add_dv(ws, c, ("dec", 0, 0.2, "Bitte einen Prozentsatz zwischen 0 % und 20 % eingeben."), None)
        unsicher = land == "Bremen"
        m.text(ws, r, 5, "Landesgesetze zu §11 Abs. 1 GrEStG; KiSt 8 % BW/BY, sonst 9 %"
               + (" – Bremen 5,5 % seit 1.7.2025, Quellen widersprüchlich" if unsicher else ""),
               border=True, size=9, wrap=True)
        m.text(ws, r, 6, "10/2026", border=True, size=9)
        m.text(ws, r, 7, "PRÜFEN" if unsicher else "sicher", border=True, bold=unsicher,
               fillc=C_ORANGE if unsicher else None, size=9, align="center")
        r += 1
    m.define("Liste_Bundesland", ws, f"$B${r0}:$B${r-1}")
    m.define("Tab_GrESt", ws, f"$C${r0}:$C${r-1}")
    m.define("Tab_KiSt", ws, f"$D${r0}:$D${r-1}")
    for n in ("Liste_Bundesland", "Tab_GrESt", "Tab_KiSt"):
        m.leaf(n)

    sec("Ampel: Schwellwerte (grün = mindestens / höchstens …)")
    p("P_Amp_IRR_gr", "IRR nach Steuern: grün ab", 0.06, F_PCT, "p.a.", "Eigene Annahme – an persönliche Renditeerwartung anpassen", sicher=False)
    p("P_Amp_IRR_ge", "IRR nach Steuern: gelb ab", 0.04, F_PCT, "p.a.", "Eigene Annahme", sicher=False)
    p("P_Amp_Vort_gr", "Vermögensvorteil ggü. Alternative: grün ab (in % des EK)", 0.0, F_PCT, "", "Eigene Annahme", sicher=False)
    p("P_Amp_Vort_ge", "Vermögensvorteil: gelb ab", -0.10, F_PCT, "", "Eigene Annahme", sicher=False)
    p("P_Amp_CF_gr", "Cashflow n. St. Jahr 1: grün ab", -100, F_EUR, "€/Monat", "Eigene Annahme", sicher=False)
    p("P_Amp_CF_ge", "Cashflow n. St. Jahr 1: gelb ab", -300, F_EUR, "€/Monat", "Eigene Annahme", sicher=False)
    p("P_Amp_KPF_gr", "Kaufpreisfaktor: grün bis", 25, F_NUM1, "x Jahresmiete", "Marktübliche Faustregel", sicher=False)
    p("P_Amp_KPF_ge", "Kaufpreisfaktor: gelb bis", 30, F_NUM1, "x Jahresmiete", "Marktübliche Faustregel", sicher=False)
    p("P_Amp_DSCR_gr", "Kapitaldienstdeckung: grün ab", 1.2, F_NUM2, "", "Bankübliche Mindestgröße", sicher=False)
    p("P_Amp_DSCR_ge", "Kapitaldienstdeckung: gelb ab", 1.0, F_NUM2, "", "Bankübliche Mindestgröße", sicher=False)
    p("P_Amp_BQKA_gr", "Belastung Kapitalanlage (Zuzahlung/Netto): grün bis", 0.10, F_PCT, "", "Eigene Annahme", sicher=False)
    p("P_Amp_BQKA_ge", "Belastung Kapitalanlage: gelb bis", 0.20, F_PCT, "", "Eigene Annahme", sicher=False)
    p("P_Amp_BQEN_gr", "Wohnkostenquote Eigennutzung: grün bis", 0.35, F_PCT, "", "Bankübliche Faustregel (35–40 %)", sicher=False)
    p("P_Amp_BQEN_ge", "Wohnkostenquote Eigennutzung: gelb bis", 0.45, F_PCT, "", "Eigene Annahme", sicher=False)
    p("P_Amp_Score_gr", "Fazit 'lohnt sich' ab Punktzahl (0–2)", 1.5, F_NUM2, "Punkte", "Eigene Annahme", sicher=False)
    p("P_Amp_Score_ge", "Fazit 'grenzwertig' ab Punktzahl", 0.9, F_NUM2, "Punkte", "Eigene Annahme", sicher=False)
    p("P_Amp_W_IRR", "Gewicht IRR", 2, F_NUM1, "", "Eigene Annahme", sicher=False)
    p("P_Amp_W_Vort", "Gewicht Vermögensvorteil", 3, F_NUM1, "", "Eigene Annahme", sicher=False)
    p("P_Amp_W_CF", "Gewicht Cashflow", 1, F_NUM1, "", "Eigene Annahme", sicher=False)
    p("P_Amp_W_KPF", "Gewicht Kaufpreisfaktor", 1, F_NUM1, "", "Eigene Annahme", sicher=False)
    p("P_Amp_W_DSCR", "Gewicht Kapitaldienstdeckung", 1, F_NUM1, "", "Eigene Annahme", sicher=False)
    p("P_Amp_W_BQ", "Gewicht Belastungsquote", 2, F_NUM1, "", "Eigene Annahme", sicher=False)

    # Auswahllisten
    r += 1
    m.section(ws, r, "Auswahllisten (nicht ändern – werden von Dropdowns verwendet)")
    r += 1
    lists = [
        ("Liste_JaNein", [JA, NEIN]), ("Liste_Modus", [KA, EN]),
        ("Liste_Szenario", ["Base", "Best", "Worst"]), ("Liste_Objekttyp", [ETW, HAUS, MFH]),
        ("Liste_Veranlagung", [EINZ, ZUS]), ("Liste_Steuermodus", [TARIF, DIREKT]),
        ("Liste_Aufteilung", [AUF_AH, AUF_MAN]), ("Liste_RenovArt", [ERH, HK]),
        ("Liste_InstModus", [INST_P, INST_M]), ("Liste_NBModus", [NB_ZIEL, NB_FIX]),
        ("Liste_NBVerw", [NB_FOLGE, NB_MOD, NB_ETF, NB_KONSUM]),
        ("Liste_Denkmal", [DENK_NEIN, DENK_7I, DENK_7H]), ("Liste_Kennzahl", KENNZAHLEN),
    ]
    for name, vals in lists:
        m.text(ws, r, 2, name, border=True, color="7F7F7F", size=9)
        for i, v in enumerate(vals):
            m.text(ws, r + i, 3, v, border=True, size=9)
        m.define(name, ws, f"$C${r}:$C${r + len(vals) - 1}")
        m.leaf(name)
        r += len(vals)
    m.protect(ws)
    ws.freeze_panes = "B4"


# ---------------------------------------------------------------------------
def build_eingaben(m):
    ws = m.sheet("ein", "Eingaben")
    m.title(ws, "Eingaben – Objekt, Finanzierung, Situation",
            "Nur die gelben Zellen ausfüllen. Die Werte sind ein Beispielfall – einfach mit deinen Zahlen überschreiben. "
            "Beim Ändern erscheint ein Hinweis; ungültige Werte werden abgewiesen.")
    for col, w in zip("BCDEFG", (44, 17, 17, 17, 17, 80)):
        ws.column_dimensions[col].width = w
    r = 4
    D = ("dec", 0, 100000000, "Bitte einen Betrag ≥ 0 eingeben (keine negativen Werte).")
    PCT = ("dec", 0, 1, "Bitte einen Prozentsatz zwischen 0 % und 100 % eingeben.")
    RATE = ("dec", 0, 0.2, "Bitte einen Zinssatz zwischen 0 % und 20 % eingeben.")
    GROW = ("dec", -0.1, 0.2, "Bitte eine Veränderung zwischen −10 % und +20 % p.a. eingeben.")

    def I(name, label, val, fmt, spec, hint, unit=""):
        nonlocal r
        m.inp(ws, r, name, label, val, fmt, spec, hint, unit=unit)
        r += 1

    def S(t):
        nonlocal r
        r += 1
        m.section(ws, r, t)
        m.subhead(ws, r + 1, ["Eingabe", "Wert", "Einheit", "", "", "Erklärung"])
        r += 2

    S("A  Grundeinstellungen")
    I("Modus", "Nutzung", KA, None, ("list", "Liste_Modus"),
      "Kapitalanlage = Vermietung. Eigennutzung = selbst bewohnen (optional mit Teilvermietung, Abschnitt G).")
    I("Szenario", "Aktives Szenario", "Base", None, ("list", "Liste_Szenario"),
      "Base = deine Eingaben. Best/Worst verändern Zins, Miete, Leerstand, Wert und Instandhaltung laut Blatt 'Szenarien & Sensitivität'.")
    I("Startjahr", "Kaufjahr (Übergabe)", 2026, F_INT, ("int", 2000, 2100, "Bitte ein Jahr zwischen 2000 und 2100 eingeben."),
      "Jahr 1 = erste 12 Monate nach Übergabe. Steuerlich wird jedes Modelljahr vereinfacht wie ein Kalenderjahr behandelt.", "Jahr")
    I("Horizont", "Betrachtungszeitraum / Verkauf nach", 30, F_INT, ("int", 1, T_MAX, f"Bitte 1 bis {T_MAX} Jahre eingeben."),
      "Am Ende wird ein (fiktiver) Verkauf gerechnet: IRR, Vermögen, Spekulationsteuer. ≥ 10 Jahre = steuerfreier Verkauf (§23).", "Jahre")

    S("B  Objekt")
    I("Objekttyp", "Objekttyp", ETW, None, ("list", "Liste_Objekttyp"), "Steuert Kaufpreisaufteilung (Miteigentumsanteil) und Instandhaltungsvorschlag.")
    I("Bundesland", "Bundesland des Objekts", "Baden-Württemberg", None, ("list", "Liste_Bundesland"), "Bestimmt den Grunderwerbsteuersatz (Blatt Parameter).")
    I("Kaufpreis", "Kaufpreis laut Vertrag (inkl. Inventar)", 318000, F_EUR, ("dec", 1000, 100000000, "Bitte einen Kaufpreis ≥ 1.000 € eingeben."),
      "Gesamtkaufpreis inkl. separat ausgewiesenem Inventar.", "€")
    I("Inventar", "davon Inventar (z. B. Einbauküche)", 8000, F_EUR, D,
      "Im Vertrag separat ausgewiesen und realistisch bewertet: keine Grunderwerbsteuer darauf, eigene AfA (Küche 10 J.).", "€")
    I("Inventar_ND", "Nutzungsdauer Inventar", 10, F_INT, ("int", 1, 30, "Bitte 1 bis 30 Jahre eingeben."), "Einbauküche: 10 Jahre (BFH IX R 14/15).", "Jahre")
    I("Wohnflaeche", "Wohnfläche", 70, F_NUM1, ("dec", 10, 100000, "Bitte eine Wohnfläche ≥ 10 m² eingeben."), "Nach Wohnflächenverordnung.", "m²")
    I("Grundstueck", "Grundstücksfläche (gesamtes Grundstück)", 1200, F_NUM, D, "Bei ETW: Fläche des gesamten Grundstücks der WEG (Grundbuch).", "m²")
    I("MEA_Anteil", "Miteigentumsanteil (nur ETW)", 0.05, F_PCT2, PCT, "Z. B. 50/1.000 = 5 %. Bei Haus/MFH wird automatisch 100 % angesetzt.", "%")
    I("Baujahr", "Baujahr / Jahr der Fertigstellung", 1995, F_INT, ("int", 1500, 2100, "Bitte ein Baujahr zwischen 1500 und 2100 eingeben."),
      "Bestimmt den AfA-Satz automatisch: vor 1925 2,5 %, 1925–2022 2 %, ab 2023 3 %.", "Jahr")
    I("Zustand", "Zustand", ZUSTAENDE[2], None, ("list", "Liste_Zustand"), "Wirkt auf Restnutzungsdauer (Kaufpreisaufteilung) und Instandhaltungsvorschlag.")
    I("Standardstufe", "Ausstattungsstandard (1 einfach – 5 gehoben)", 3, F_INT, ("int", 1, 5, "Bitte eine Stufe von 1 bis 5 eingeben."), "Standardstufe nach NHK 2010 (für die Kaufpreisaufteilung).")
    I("Bodenrichtwert", "Bodenrichtwert", 650, F_EUR, D, "Aus BORIS (Bodenrichtwert-Informationssystem) des Landes.", "€/m²")
    I("Aufteilung_Modus", "Kaufpreisaufteilung Gebäude/Boden", AUF_AH, None, ("list", "Liste_Aufteilung"),
      "Arbeitshilfe = vereinfachte BMF-Logik. Manuell = Aufteilung aus Kaufvertrag oder Gutachten (hat Vorrang, wenn plausibel).")
    I("Gebaeudeanteil_man", "Gebäudeanteil manuell", 0.75, F_PCT, PCT, "Nur bei 'Manuell': Anteil des Gebäudes am Kaufpreis (ohne Inventar).", "%")
    I("Renovierung", "Renovierung/Modernisierung zum Start", 10000, F_EUR, D, "Brutto inkl. USt, wird mitfinanziert. Bei Eigennutzung Basis für §35a/§35c.", "€")
    I("Renov_Art", "Art der Start-Renovierung", ERH, None, ("list", "Liste_RenovArt"),
      "Erhaltung = sofort abziehbar (sofern 15-%-Grenze nicht überschritten). Herstellung = Erweiterung/Standardhebung → AfA.")
    I("Renov_Lohn", "Lohnanteil der Renovierung", 0.40, F_PCT, PCT, "Für §35a (nur Eigennutzung).", "%")
    I("Renov_Energ", "davon energetische Maßnahmen", 0.0, F_PCT, PCT, "Für §35c (nur Eigennutzung, Gebäude > 10 Jahre, Fachunternehmer-Bescheinigung).", "%")
    I("Energ_gefoerdert", "Energet. Maßnahme mit KfW/BAFA gefördert?", NEIN, None, ("list", "Liste_JaNein"), "Wenn Ja, ist §35c ausgeschlossen.")
    I("Renov_Wert", "Wertwirksamer Anteil der Renovierung", 0.5, F_PCT, PCT, "Wie viel der Renovierung den Marktwert erhöht.", "%")

    S("C  Kaufnebenkosten")
    I("Notar_Satz", "Notar", 0.015, F_PCT2, ("dec", 0, 0.05, "Bitte 0 % bis 5 % eingeben."), "Übliche Spanne 1,0–2,0 % inkl. Grundschuldbestellung.", "% v. KP")
    I("Grundbuch_Satz", "Grundbuchamt", 0.005, F_PCT2, ("dec", 0, 0.05, "Bitte 0 % bis 5 % eingeben."), "Übliche Spanne 0,3–0,6 %.", "% v. KP")
    I("Grundschuld_Anteil", "Anteil Notar/Grundbuch für Grundschuld", 0.30, F_PCT, PCT,
      "Finanzierungskosten (Grundschuld) sind bei Vermietung sofort als Werbungskosten abziehbar, der Rest erhöht die Anschaffungskosten.", "%")
    I("Makler_Satz", "Makler (Käuferanteil inkl. USt)", 0.0357, F_PCT2, ("dec", 0, 0.1, "Bitte 0 % bis 10 % eingeben."), "Bei Wohnungen/EFH max. Hälfte der Gesamtprovision (§656c BGB).", "% v. KP")
    I("NK_Sonstige", "Sonstige Erwerbskosten", 0, F_EUR, D, "Z. B. Wertgutachten, Grundbuchauszüge.", "€")
    I("RND_Gutachten", "Gutachten kürzere Restnutzungsdauer?", NEIN, None, ("list", "Liste_JaNein"),
      "§7 Abs. 4 S. 2 EStG: kürzere tatsächliche Nutzungsdauer per Gutachten → höhere AfA (nur wenn > Regel-AfA).")
    I("RND_Jahre", "Restnutzungsdauer laut Gutachten", 40, F_INT, ("int", 1, 100, "Bitte 1 bis 100 Jahre eingeben."), "AfA = 1 / Restnutzungsdauer.", "Jahre")
    I("RND_Kosten", "Kosten des RND-Gutachtens", 1000, F_EUR, D, "Als Werbungskosten im Jahr 1 angesetzt.", "€")

    S("D  Eigenkapital & Finanzierung")
    I("Eigenkapital", "Verfügbares Eigenkapital", 75000, F_EUR, D, "Bares Eigenkapital. Darlehen 1 deckt automatisch den Rest.", "€")
    I("Riester", "Wohn-Riester-Entnahme (nur Eigennutzung)", 0, F_EUR, D,
      "Geförderte Riester-Kapital als EK (§92a EStG). Achtung: Wohnförderkonto wird im Alter nachgelagert besteuert.", "€")
    r += 1
    m.subhead(ws, r, ["Darlehen", "Darlehen 1 (Restbedarf)", "Darlehen 2", "Darlehen 3", "Darlehen 4", "Erklärung"])
    r += 1
    loans = [
        # Feld, Label, fmt, spec, hint, [K1..K4]
        ("Name", "Bezeichnung", None, ("text",), "Frei wählbar.",
         ["Bank Hauptdarlehen", "Bank Tranche 2", "Bausparen/Sonstiges", "KfW-Förderdarlehen"]),
        ("Betrag", "Darlehensbetrag", F_EUR, D, "Darlehen 1 wird automatisch berechnet (Gesamtkosten − EK − Darlehen 2–4).", [None, 60000, 0, 0]),
        ("Disagio", "Disagio", F_PCT2, ("dec", 0, 0.1, "Bitte 0 % bis 10 % eingeben."),
         "Abschlag auf die Auszahlung; bis 5 % bei ≥ 5 J. Zinsbindung sofort abziehbar, sonst über Zinsbindung verteilt.", [0, 0, 0, 0]),
        ("Zins", "Sollzins p.a.", F_PCT2, RATE, "Gebundener Sollzins.", [0.037, 0.039, 0.03, 0.025]),
        ("Tilgung", "Anfängliche Tilgung p.a.", F_PCT2, ("dec", 0, 0.5, "Bitte 0 % bis 50 % eingeben."), "0 % = endfälliges Darlehen.", [0.02, 0.02, 0.03, 0.02]),
        ("ZB", "Zinsbindung", F_INT, ("int", 1, 40, "Bitte 1 bis 40 Jahre eingeben."), "Danach gilt der Anschlusszins.", [10, 15, 10, 10]),
        ("Anschluss", "Anschlusszins (Base) p.a.", F_PCT2, RATE, "Annahme für die Zeit nach der Zinsbindung; Best/Worst verschieben ihn.", [0.045, 0.045, 0.045, 0.045]),
        ("AT", "Anschlusstilgung p.a.", F_PCT2, ("dec", 0, 0.5, "Bitte 0 % bis 50 % eingeben."), "Rate nach Zinsbindung = Restschuld × (Anschlusszins + Anschlusstilgung).", [0.025, 0.025, 0.03, 0.025]),
        ("TF", "Tilgungsfreie Anlaufjahre", F_INT, ("int", 0, 10, "Bitte 0 bis 10 Jahre eingeben."), "Typisch bei KfW (nur Zinsen).", [0, 0, 0, 1]),
        ("ST", "Sondertilgung p.a.", F_EUR, D, "Am Jahresende, sofern Restschuld vorhanden.", [0, 0, 0, 0]),
        ("ST_ab", "Sondertilgung ab Jahr", F_INT, ("int", 1, 40, "Bitte 1 bis 40 eingeben."), "", [1, 1, 1, 1]),
        ("ST_bis", "Sondertilgung bis Jahr", F_INT, ("int", 1, 40, "Bitte 1 bis 40 eingeben."), "", [10, 15, 10, 10]),
        ("TZ", "Tilgungszuschuss (Förderung)", F_PCT, PCT, "In % des Darlehens (z. B. KfW). Mindert die Restschuld; bei Vermietung steuerlich als Einnahme angesetzt (vereinfacht).", [0, 0, 0, 0]),
        ("TZ_Jahr", "Tilgungszuschuss gutgeschrieben in Jahr", F_INT, ("int", 1, 40, "Bitte 1 bis 40 eingeben."), "", [1, 1, 1, 2]),
        ("BZ", "Bereitstellungszins pro Monat", F_PCT2, ("dec", 0, 0.02, "Bitte 0 % bis 2 % pro Monat eingeben."), "Üblich 0,25 % pro Monat.", [0.0025, 0.0025, 0.0025, 0.0025]),
        ("BZ_Monate", "Monate mit Bereitstellungszins", F_INT, ("int", 0, 36, "Bitte 0 bis 36 Monate eingeben."),
         "Monate nach der bereitstellungszinsfreien Zeit bis zur Auszahlung (bei Bestandskauf meist 0).", [0, 0, 0, 0]),
    ]
    for field, label, fmt, spec, hint, vals in loans:
        m.label(ws, r, label, None, hint)
        for k in range(1, 5):
            name = f"Kredit{k}_{field}"
            if field == "Betrag" and k == 1:
                c = ws.cell(row=r, column=3, value="=Kredit1_Betrag")
                m.style(c, fmt=F_EUR, fillc=C_GREY, italic=True)
                continue
            m.inp(ws, r, name, label, vals[k - 1], fmt, spec, hint, col=2 + k, write_label=False)
        r += 1

    S("E  Laufende Einnahmen & Kosten")
    I("Kaltmiete", "Kaltmiete (Kapitalanlage)", 840, F_EUR, D, "Vereinbarte Nettokaltmiete pro Monat.", "€/Monat")
    I("Vergleichsmiete", "Ortsübliche Vergleichsmiete (kalt)", 840, F_EUR, D, "Mietspiegel. Für die 66-%-Prüfung bei Vermietung an Angehörige.", "€/Monat")
    I("Angehoerige", "Vermietung an Angehörige / verbilligt?", NEIN, None, ("list", "Liste_JaNein"), "Prüft die 66-%/50-%-Grenze nach §21 Abs. 2 EStG.")
    I("Mietsteigerung", "Mietsteigerung p.a.", 0.02, F_PCT, GROW, "Kappungsgrenze/Mietpreisbremse beachten (in angespannten Märkten 15 %/3 J.).", "% p.a.")
    I("Leerstand", "Leerstand / Mietausfall", 0.03, F_PCT, PCT, "Anteil der Jahresmiete. 2–4 % sind üblich.", "%")
    I("HG_umlage", "Hausgeld umlagefähig", 190, F_EUR, D, "Betriebskosten, die der Mieter über Nebenkosten zahlt (z. B. Wasser, Müll, Hausmeister).", "€/Monat")
    I("HG_nicht_umlage", "Hausgeld nicht umlagefähig (ohne Rücklage)", 45, F_EUR, D, "WEG-Verwaltung, Kontoführung WEG, Reparaturen Gemeinschaft.", "€/Monat")
    I("HG_Ruecklage", "Zuführung Erhaltungsrücklage (WEG)", 70, F_EUR, D, "Teil des Hausgelds. Faustregel ETW ca. 0,80–1,50 €/m² monatlich.", "€/Monat")
    I("Ruecklage_sofort", "Rücklage sofort als Werbungskosten ansetzen?", JA, None, ("list", "Liste_JaNein"),
      "Steuerlich erst bei Verwendung abziehbar (BFH). 'Ja' = Vereinfachung (Verwendung ≈ Zuführung). 'Nein' = konservativ.")
    I("Grundsteuer", "Grundsteuer", 300, F_EUR, D, "Bei Vermietung umlagefähig.", "€/Jahr")
    I("Vers_umlage", "Versicherungen umlagefähig", 0, F_EUR, D, "Wohngebäude/Haftpflicht, falls nicht im Hausgeld.", "€/Jahr")
    I("Vers_nicht_umlage", "Versicherungen nicht umlagefähig", 100, F_EUR, D, "Z. B. Mietausfall-, Vermieterrechtsschutzversicherung.", "€/Jahr")
    I("Mietverwaltung", "Mietverwaltung (Sondereigentum)", 0, F_EUR, D, "Nicht umlagefähig. Typisch 25–35 € pro Einheit und Monat.", "€/Monat")
    I("Inst_Modus", "Eigene Instandhaltungsreserve", INST_P, None, ("list", "Liste_InstModus"),
      "Peters'sche Formel: Herstellkosten/m² × 1,5 / 80; bei ETW nur Anteil Sondereigentum.")
    I("Peters_HK", "Herstellkosten-Ansatz für Peters'sche Formel", 3000, F_EUR, D, "Heutige Baukosten je m² Wohnfläche.", "€/m²")
    I("Inst_manuell", "Instandhaltungsreserve manuell", 15, F_EUR2, D, "Nur bei 'Manuell'.", "€/m²·Jahr")
    I("Erh_Betrag", "Geplanter größerer Erhaltungsaufwand", 12000, F_EUR, D, "Z. B. neues Bad. Wird aus laufendem Cashflow bezahlt.", "€")
    I("Erh_Jahr", "… im Jahr", 12, F_INT, ("int", 1, 40, "Bitte 1 bis 40 eingeben."),
      "Achtung: in den Jahren 1–3 zählt er zur 15-%-Grenze (anschaffungsnahe Herstellungskosten).", "Jahr")
    I("Verteilung_82b", "Erhaltungsaufwand verteilen auf", 1, F_INT, ("int", 1, 5, "Bitte 1 bis 5 Jahre eingeben."),
      "§82b EStDV: größerer Erhaltungsaufwand bei vermieteten Wohngebäuden auf 2–5 Jahre verteilbar (glättet die Progression). 1 = sofort.", "Jahre")
    I("Fahrt_km", "Fahrten zum Objekt", 200, F_NUM, D, "Gefahrene Kilometer pro Jahr (Besichtigungen, Übergaben).", "km/Jahr")
    I("Kontofuehrung", "Kontoführung Mietkonto", 16, F_EUR, D, "Pauschal 16 € ohne Nachweis üblich.", "€/Jahr")
    I("Steuerberatung", "Steuerberatung (Anteil Vermietung)", 250, F_EUR, D, "Als Werbungskosten abziehbar.", "€/Jahr")
    I("Sonstige_WK", "Sonstige Werbungskosten", 100, F_EUR, D, "Porto, Telefon, Inserate, Fachliteratur, Haus & Grund-Beitrag …", "€/Jahr")

    S("F  Wertentwicklung, Verkauf & Alternative")
    I("Wertsteigerung", "Wertsteigerung Immobilie p.a.", 0.015, F_PCT, GROW, "Langfristig nominal oft 1–2,5 % p.a.; stark lageabhängig.", "% p.a.")
    I("Inflation", "Inflation / Kostensteigerung p.a.", 0.02, F_PCT, GROW, "Steigerung von Hausgeld, Instandhaltung und sonstigen Kosten.", "% p.a.")
    I("Verkaufskosten", "Verkaufskosten", 0.035, F_PCT, ("dec", 0, 0.2, "Bitte 0 % bis 20 % eingeben."), "Makler Verkäuferanteil, Energieausweis, Löschung Grundschuld.", "% v. Wert")
    I("VFE_Satz", "Vorfälligkeitsentschädigung (Schätzung)", 0.03, F_PCT, ("dec", 0, 0.3, "Bitte 0 % bis 30 % eingeben."),
      "Nur bei Verkauf innerhalb Zinsbindung und vor Ablauf von 10 Jahren (§489 BGB).", "% d. RS")
    I("Objektverkaeufe", "Objektverkäufe in 5 Jahren nach Kauf (inkl. diesem)", 1, F_INT, ("int", 0, 50, "Bitte 0 bis 50 eingeben."),
      "Mehr als 3 Objekte binnen 5 Jahren → Gefahr gewerblicher Grundstückshandel (Drei-Objekt-Grenze).", "Anzahl")
    I("ETF_Rendite", "Rendite Alternative (ETF) p.a. vor Steuern", 0.06, F_PCT, ("dec", -0.2, 0.3, "Bitte −20 % bis +30 % eingeben."),
      "Breiter Aktien-ETF: langfristig ca. 5–7 % p.a. nominal (nach Kosten). Teilfreistellung/Abgeltungsteuer auf Blatt Parameter.", "% p.a.")

    S("G  Eigennutzung (nur relevant bei Nutzung = Eigennutzung)")
    I("Miete_Alternativ", "Kaltmiete einer vergleichbaren Mietwohnung", 840, F_EUR, D,
      "Für den Vergleich Kaufen vs. Mieten (nur selbst genutzter Teil).", "€/Monat")
    I("Teilverm_Anteil", "Teilvermietung: vermieteter Flächenanteil", 0, F_PCT, PCT, "Z. B. Einliegerwohnung. Kosten/AfA werden anteilig steuerlich berücksichtigt.", "%")
    I("Teilverm_Miete", "Teilvermietung: Kaltmiete", 0, F_EUR, D, "Kaltmiete des vermieteten Teils.", "€/Monat")
    I("Teilverm_Zuordnung", "Darlehen gezielt dem vermieteten Teil zuordnen?", NEIN, None, ("list", "Liste_JaNein"),
      "Legaler Hebel (BFH IX R 18/97): separates Darlehen bezahlt nachweislich den vermieteten Teil → Zinsen dort voll abziehbar.")
    I("Handwerker_Lohn", "Handwerkerleistungen (Lohnanteil) p.a.", 500, F_EUR, D, "§35a Abs. 3: 20 % der Lohnkosten, max. 1.200 € Steuerermäßigung.", "€/Jahr")
    I("HHDL", "Haushaltsnahe Dienstleistungen (Lohn) p.a.", 300, F_EUR, D, "§35a Abs. 2: z. B. Hausmeister/Treppenreinigung aus der Hausgeldabrechnung, 20 %, max. 4.000 €.", "€/Jahr")
    I("Netto_Einkommen", "Netto-Haushaltseinkommen", 3800, F_EUR, D, "Für Belastungsquote / Wohnkostenquote in der Ampel.", "€/Monat")

    S("H  Steuerliche Situation")
    I("Steuer_Modus", "Berechnung des Steuereffekts", TARIF, None, ("list", "Liste_Steuermodus"),
      "Empfohlen: aus zvE berechnen (exakte Tarifdifferenz inkl. Progression, Soli, KiSt). Alternativ fester Grenzsteuersatz.")
    I("zvE", "Zu versteuerndes Einkommen (ohne dieses Objekt)", 70000, F_EUR, D, "Aus dem letzten Steuerbescheid. Bei Zusammenveranlagung gemeinsames zvE.", "€/Jahr")
    I("zvE_Steigerung", "Steigerung zvE p.a.", 0.0, F_PCT, GROW, "Tarif bleibt auf Stand 2026 (keine Indexierung) – daher eher 0 % ansetzen.", "% p.a.")
    I("Veranlagung", "Veranlagung", EINZ, None, ("list", "Liste_Veranlagung"), "Zusammenveranlagung = Splittingtarif.")
    I("Grenzsteuersatz_direkt", "Grenzsteuersatz (nur Modus 'direkt')", 0.42, F_PCT, ("dec", 0, 0.45, "Bitte 0 % bis 45 % eingeben."), "Ohne Soli/KiSt.", "%")
    I("Soli_direkt", "Soli fällig? (nur Modus 'direkt')", NEIN, None, ("list", "Liste_JaNein"), "Im Modus 'zvE' wird der Soli automatisch berechnet.")
    I("Kirchensteuer", "Kirchensteuerpflichtig?", NEIN, None, ("list", "Liste_JaNein"), "Satz nach Wohnsitz-Bundesland (BW/BY 8 %, sonst 9 %).")
    I("Wohnsitz_Land", "Wohnsitz-Bundesland", "Baden-Württemberg", None, ("list", "Liste_Bundesland"), "Für den Kirchensteuersatz.")

    S("I  Neubau / Denkmal (Sonderabschreibungen)")
    I("Erwerb_FJ", "Erwerb bis Ende des Fertigstellungsjahres?", NEIN, None, ("list", "Liste_JaNein"),
      "Voraussetzung für degressive AfA und §7b beim Kauf (sonst nur für Bauherren).")
    I("Degr_Wahl", "Degressive AfA (5 %) nutzen?", NEIN, None, ("list", "Liste_JaNein"), "Nur wenn berechtigt (Baubeginn 1.10.2023–30.9.2029). Wechsel zu linear erfolgt automatisch, sobald günstiger.")
    I("Baubeginn", "Baubeginn (Bauantrag/Bauanzeige)", dt.date(1994, 3, 1), F_DATE, ("date",), "Für die degressive AfA maßgeblich.", "Datum")
    I("Sonder7b_Wahl", "Sonder-AfA §7b nutzen?", NEIN, None, ("list", "Liste_JaNein"), "5 % p.a. für 4 Jahre zusätzlich; nur bei Neubau-Mietwohnungen mit EH40/QNG.")
    I("Bauantrag", "Datum Bauantrag", dt.date(1993, 6, 1), F_DATE, ("date",), "§7b: Bauantrag 1.1.2023–30.9.2029.", "Datum")
    I("Effizienz_EH40", "Effizienzhaus 40 mit QNG-Siegel?", NEIN, None, ("list", "Liste_JaNein"), "§7b-Voraussetzung für Bauanträge ab 2023. Wohnung muss 10 Jahre vermietet werden.")
    I("Denkmal_Status", "Denkmal / Sanierungsgebiet", DENK_NEIN, None, ("list", "Liste_Denkmal"),
      "§7i/§7h (Vermietung: 8×9 % + 4×7 %), §10f (Eigennutzung: 10×9 %). Bescheinigung der Denkmalbehörde/Gemeinde nötig.")
    I("Denkmal_Betrag", "Begünstigte Sanierungskosten (im Kaufpreis)", 0, F_EUR, D,
      "Anteil des Kaufpreises, der laut Bauträger/Bescheinigung auf begünstigte Maßnahmen entfällt (Bauträgermodell).", "€")

    S("J  Nachbeleihung (später erneut belasten)")
    I("NB_Aktiv", "Nachbeleihung einplanen?", NEIN, None, ("list", "Liste_JaNein"),
      "Nach einigen Jahren neues Darlehen auf die (wertgestiegene) Immobilie aufnehmen, z. B. als EK für das nächste Objekt.")
    I("NB_Jahr", "… zu Beginn von Jahr", 11, F_INT, ("int", 1, 40, "Bitte 1 bis 40 eingeben."), "Sinnvoll nach Ende der Zinsbindung (keine Vorfälligkeit).", "Jahr")
    I("NB_Modus", "Höhe der Nachbeleihung", NB_ZIEL, None, ("list", "Liste_NBModus"), "Fester Betrag oder Aufstockung bis zum Ziel-Beleihungsauslauf.")
    I("NB_Betrag", "Fester Betrag", 50000, F_EUR, D, "Nur bei 'Fester Betrag'.", "€")
    I("NB_Ziel", "Ziel-Beleihungsauslauf (Darlehen/Marktwert)", 0.60, F_PCT, PCT, "Banken beleihen meist bis 60–80 % des Beleihungswerts.", "%")
    I("NB_Zins", "Zins Nachbeleihung p.a.", 0.045, F_PCT2, RATE, "Best/Worst verschieben ihn wie den Anschlusszins.", "% p.a.")
    I("NB_Tilgung", "Tilgung Nachbeleihung p.a.", 0.02, F_PCT2, ("dec", 0, 0.5, "Bitte 0 % bis 50 % eingeben."), "", "% p.a.")
    I("NB_Verwendung", "Verwendung des Geldes", NB_FOLGE, None, ("list", "Liste_NBVerw"),
      "Steuerlich entscheidend: Zinsen nur abziehbar bei Verwendung für Einkünfte aus V+V (Modernisierung vermietet, weiteres Mietobjekt). ETF: §20 Abs. 9 (nicht abziehbar). Konsum: nicht abziehbar.")
    I("NB_Mod_Art", "Bei Modernisierung: Art", ERH, None, ("list", "Liste_RenovArt"), "Erhaltung (sofort/§82b) oder Herstellung (AfA).")
    I("NB_Wertanteil", "Wertwirksamer Anteil Modernisierung", 0.7, F_PCT, PCT, "", "%")
    I("NB_Folge_Rendite", "Rendite des Folgeobjekts auf das EK p.a. (netto)", 0.05, F_PCT, ("dec", -0.2, 0.3, "Bitte −20 % bis +30 % eingeben."),
      "Vereinfacht: Wertzuwachs des eingesetzten EK im Folgeobjekt nach Steuern.", "% p.a.")
    ws.freeze_panes = "C4"
    m.protect(ws)


# ---------------------------------------------------------------------------
def build_finanzierung(m):
    ws = m.sheet("fin", "Kaufnebenkosten & Finanzierung")
    m.title(ws, "Kaufnebenkosten, Kaufpreisaufteilung & Finanzierung",
            "Rechenblatt – keine Eingaben. Werte stammen vom Blatt 'Eingaben' und 'Parameter'.")
    for col, w in zip("BCDEFGHIJKLMNO", (44, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["G"].width = 60
    r = 4

    def C(name, label, tpl, fmt, hint=None, key=False, unit=None):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, hint, key=key, unit=unit)
        r += 1

    m.section(ws, r, "Kaufnebenkosten"); r += 1
    C("KP_Immobilie", "Kaufpreis Immobilie (ohne Inventar)", "{Kaufpreis}-{Inventar}", F_EUR, "Bemessungsgrundlage Grunderwerbsteuer.")
    C("GrESt_Satz", "Grunderwerbsteuersatz", "INDEX({Tab_GrESt},MATCH({Bundesland},{Liste_Bundesland},0))", F_PCT, "Nach Bundesland (Blatt Parameter).")
    C("GrESt", "Grunderwerbsteuer", "{KP_Immobilie}*{GrESt_Satz}", F_EUR)
    C("Notar", "Notar", "{Kaufpreis}*{Notar_Satz}", F_EUR)
    C("Grundbuch", "Grundbuch", "{Kaufpreis}*{Grundbuch_Satz}", F_EUR)
    C("Makler", "Makler", "{Kaufpreis}*{Makler_Satz}", F_EUR)
    C("NK_Finanzierung", "davon Finanzierungskosten (Grundschuld)", "({Notar}+{Grundbuch})*{Grundschuld_Anteil}", F_EUR, "Bei Vermietung sofort abziehbar.")
    C("NK_Erwerb", "Erwerbsnebenkosten (aktivierungspflichtig)", "{GrESt}+({Notar}+{Grundbuch})*(1-{Grundschuld_Anteil})+{Makler}+{NK_Sonstige}", F_EUR, "Erhöhen die Anschaffungskosten (Boden + Gebäude).")
    C("NK_Summe", "Kaufnebenkosten gesamt", "{NK_Erwerb}+{NK_Finanzierung}", F_EUR, key=True)
    C("NK_Quote", "Kaufnebenkosten in % des Kaufpreises", "{NK_Summe}/{Kaufpreis}", F_PCT)
    C("RND_Kosten_eff", "Gutachten Restnutzungsdauer", 'IF({RND_Gutachten}="Ja",{RND_Kosten},0)', F_EUR)
    C("Gesamtkosten", "Gesamtinvestition (inkl. Renovierung)", "{Kaufpreis}+{NK_Summe}+{RND_Kosten_eff}+{Renovierung}", F_EUR, key=True)

    r += 1
    m.section(ws, r, "Kaufpreisaufteilung Grund und Boden / Gebäude (vereinfachte BMF-Arbeitshilfe)"); r += 1
    C("MEA_eff", "Miteigentumsanteil", f'IF({{Objekttyp}}="{ETW}",{{MEA_Anteil}},1)', F_PCT2)
    C("Bodenwert_AH", "Bodenwert = BRW × Fläche × MEA", "{Bodenrichtwert}*{Grundstueck}*{MEA_eff}", F_EUR)
    C("Alter", "Gebäudealter im Kaufjahr", "MAX(0,{Startjahr}-{Baujahr})", F_INT, unit="Jahre")
    C("RND_AH", "Restnutzungsdauer (modifiziert)", "MAX({P_GND}*{P_RND_min},MIN({P_GND},{P_GND}-{Alter}+INDEX({Tab_Zust_Verj},MATCH({Zustand},{Liste_Zustand},0))))", F_NUM1, unit="Jahre")
    C("BGF_m2", "Brutto-Grundfläche (geschätzt)", "{Wohnflaeche}*INDEX({Tab_BGF},MATCH({Objekttyp},{Tab_Objekttyp},0))", F_M2)
    C("NHK_m2", "NHK 2010 × Standard × Baupreisindex", "INDEX({Tab_NHK},MATCH({Objekttyp},{Tab_Objekttyp},0))*INDEX({Tab_Standard},{Standardstufe})*{P_BPI}/100", F_EUR, unit="€/m² BGF")
    C("Gebaeudewert_AH", "Gebäudesachwert (alterswertgemindert)", "{BGF_m2}*{NHK_m2}*{RND_AH}/{P_GND}", F_EUR)
    C("Gebaeudeanteil_AH", "Gebäudeanteil laut Arbeitshilfe-Logik", "IF({Gebaeudewert_AH}+{Bodenwert_AH}>0,{Gebaeudewert_AH}/({Gebaeudewert_AH}+{Bodenwert_AH}),0.8)", F_PCT,
      "Verhältnis Gebäudewert zu Gesamtwert, angewendet auf Kaufpreis + Nebenkosten.")
    C("Gebaeudeanteil", "Verwendeter Gebäudeanteil", f'IF({{Aufteilung_Modus}}="{AUF_MAN}",{{Gebaeudeanteil_man}},{{Gebaeudeanteil_AH}})', F_PCT, key=True,
      hint="Die BMF-Arbeitshilfe ist nicht bindend (BFH IX R 26/19). JStG 2026 (Entwurf §6f EStG) plant gesetzliche Regelung – prüfen.")
    C("AK_Immobilie", "Anschaffungskosten Immobilie (ohne Inventar)", "{KP_Immobilie}+{NK_Erwerb}", F_EUR)
    C("AK_Gebaeude", "davon Gebäude (AfA-Basis)", "{AK_Immobilie}*{Gebaeudeanteil}", F_EUR, key=True)
    C("AK_Boden", "davon Grund und Boden (keine AfA)", "{AK_Immobilie}-{AK_Gebaeude}", F_EUR)
    C("Startwert", "Marktwert zu Beginn (für Wertentwicklung)", "{KP_Immobilie}+{Renovierung}*{Renov_Wert}", F_EUR, "Kaufpreis ohne Inventar + wertwirksamer Teil der Renovierung.")

    r += 1
    m.section(ws, r, "Finanzierungsstruktur"); r += 1
    C("EK_verfuegbar", "Verfügbares Eigenkapital (inkl. Wohn-Riester)", f'{{Eigenkapital}}+IF({{Modus}}="{EN}",{{Riester}},0)', F_EUR)
    C("Fin_Bedarf", "Finanzierungsbedarf (Auszahlung)", "MAX(0,{Gesamtkosten}-{EK_verfuegbar})", F_EUR)
    for k in (2, 3, 4):
        C(f"Kredit{k}_Netto", f"Auszahlung Darlehen {k} (nach Disagio)", f"{{Kredit{k}_Betrag}}*(1-{{Kredit{k}_Disagio}})", F_EUR)
    C("Kredit1_Netto", "Auszahlung Darlehen 1 (Restbedarf)", "MAX(0,{Fin_Bedarf}-{Kredit2_Netto}-{Kredit3_Netto}-{Kredit4_Netto})", F_EUR)
    C("Kredit1_Betrag", "Nominalbetrag Darlehen 1", "{Kredit1_Netto}/(1-{Kredit1_Disagio})", F_EUR, key=True)
    C("Darlehen_Summe", "Darlehen gesamt (nominal)", "{Kredit1_Betrag}+{Kredit2_Betrag}+{Kredit3_Betrag}+{Kredit4_Betrag}", F_EUR, key=True)
    C("Auszahlung_Summe", "Auszahlungen gesamt", "{Kredit1_Netto}+{Kredit2_Netto}+{Kredit3_Netto}+{Kredit4_Netto}", F_EUR)
    C("EK_Einsatz", "Tatsächlich eingesetztes Eigenkapital", "{Gesamtkosten}-{Auszahlung_Summe}", F_EUR, key=True,
      hint="Startkapital für IRR und Vergleich mit der ETF-Alternative.")
    C("Disagio_Summe", "Disagio gesamt", "{Darlehen_Summe}-{Auszahlung_Summe}", F_EUR)
    C("Bereitstellung", "Bereitstellungszinsen (Jahr 1)", "+".join(f"{{Kredit{k}_Betrag}}*{{Kredit{k}_BZ}}*{{Kredit{k}_BZ_Monate}}" for k in range(1, 5))
      .replace("{Kredit1_Betrag}*", "{Kredit1_Betrag}*"), F_EUR, "Werbungskosten bei Vermietung.")
    C("Beleihungsauslauf", "Beleihungsauslauf (Darlehen / Kaufpreis)", "{Darlehen_Summe}/{Kaufpreis}", F_PCT)
    C("Fin_Hinweis", "Hinweis", f'IF({{EK_Einsatz}}<{{EK_verfuegbar}}-1,"ℹ Darlehen 2–4 übersteigen den Bedarf – Überfinanzierung, EK wird nicht voll benötigt.",IF({{Beleihungsauslauf}}>1,"⚠ Finanzierung über 100 % des Kaufpreises – höhere Zinsaufschläge wahrscheinlich.","✔ Finanzierung plausibel."))', None)

    r += 1
    m.section(ws, r, "Darlehensübersicht", ncols=7); r += 1
    m.subhead(ws, r, ["Kennzahl", "Darlehen 1", "Darlehen 2", "Darlehen 3", "Darlehen 4", "Erklärung"]); r += 1
    rows = [
        ("Bezeichnung", "{Kredit%d_Name}", None),
        ("Nominalbetrag", "{Kredit%d_Betrag}", F_EUR),
        ("Rate pro Monat Jahr 1", "INDEX([[L%d_Ann]],2)/12", F_EUR2),
        ("Zinssatz Jahr 1", "INDEX([[L%d_Zins]],2)", F_PCT2),
        ("Zinsen Jahr 1", "INDEX([[L%d_Zinsen]],2)", F_EUR),
        ("Restschuld Ende Zinsbindung", "INDEX([[L%d_RS_E]],MIN(41,{Kredit%d_ZB}+1))", F_EUR),
        ("Zinssumme über Horizont", "SUMPRODUCT(([[t]]>=1)*([[t]]<={Horizont})*[[L%d_Zinsen]])", F_EUR),
    ]
    for label, tpl, fmt in rows:
        m.label(ws, r, label)
        for k in range(1, 5):
            m.view(ws, r, 2 + k, tpl.replace("%d", str(k)), fmt)
        r += 1

    r += 1
    m.section(ws, r, "Tilgungsplan (Restschuld je Jahresende)", ncols=14); r += 1
    hdr = ["Jahr", "RS Darlehen 1", "RS Darlehen 2", "RS Darlehen 3", "RS Darlehen 4", "RS Nachbeleihung",
           "Restschuld gesamt", "Zinsen", "Tilgung inkl. Sonder", "Tilgungszuschuss", "Immobilienwert", "Beleihung (RS/Wert)"]
    m.subhead(ws, r, hdr); r += 1
    m.tilg_row0 = r
    for t in range(0, T_MAX + 1):
        c = ws.cell(row=r, column=2, value=t)
        m.style(c, fmt=F_INT, align="center")
        cols = ["L1_RS_E", "L2_RS_E", "L3_RS_E", "L4_RS_E", "L5_RS_E", "RS_Summe", "Zinsen_Summe",
                "Tilgung_Summe", "Zuschuss_Summe", "Wert"]
        for i, key in enumerate(cols):
            m.view(ws, r, 3 + i, f"INDEX([[{key}]],{t + 1})", F_EUR)
        m.view(ws, r, 13, f"IF(INDEX([[Wert]],{t + 1})>0,INDEX([[RS_Summe]],{t + 1})/INDEX([[Wert]],{t + 1}),0)", F_PCT)
        r += 1
    ws.conditional_formatting.add(f"B{m.tilg_row0}:M{r - 1}",
                                  FormulaRule(formula=[f"$B{m.tilg_row0}>Horizont"], font=Font(color="A6A6A6")))
    ws.freeze_panes = "C4"
    m.protect(ws)


# ---------------------------------------------------------------------------
def build_steuer_scalars(m):
    """Steuerliche Skalare (Blatt Steuer, oberer Teil)."""
    ws = m.sheet("st", "Steuer")
    m.title(ws, "Steuer – Vermietung & Eigennutzung",
            "Rechenblatt. Grundlage: EStG Stand 2026. Keine Steuerberatung – Ergebnisse mit Steuerberater prüfen.")
    for col, w in zip("BCDEFGHIJKLMNOPQRS", (48, 16, 12, 14, 14, 70) + (14,) * 12):
        ws.column_dimensions[col].width = w
    r = 4

    def C(name, label, tpl, fmt, hint=None, key=False, unit=None):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, hint, key=key, unit=unit)
        r += 1

    m.section(ws, r, "Persönlicher Steuersatz"); r += 1
    C("Split", "Splitting-Faktor", f'IF({{Veranlagung}}="{ZUS}",2,1)', F_INT)
    C("Soli_FG", "Soli-Freigrenze (auf ESt)", "{P_Soli_Freigrenze}*{Split}", F_EUR)
    C("KiSt_Satz_eff", "Kirchensteuersatz", f'IF({{Kirchensteuer}}="{JA}",INDEX({{Tab_KiSt}},MATCH({{Wohnsitz_Land}},{{Liste_Bundesland}},0)),0)', F_PCT)
    C("ESt_zvE", "Tarifliche ESt auf zvE (2026)", T_tpl("{zvE}"), F_EUR)
    C("ESt_zvE100", "… auf zvE + 100 €", T_tpl("{zvE}+100"), F_EUR)
    C("Steuer_zvE", "ESt + Soli + KiSt auf zvE", total_tax_tpl("{ESt_zvE}"), F_EUR)
    C("Steuer_zvE100", "… auf zvE + 100 €", total_tax_tpl("{ESt_zvE100}"), F_EUR)
    C("Grenzsteuersatz_calc", "Grenzsteuersatz inkl. Soli/KiSt (berechnet)", "({Steuer_zvE100}-{Steuer_zvE})/100", F_PCT, key=True,
      hint="So viel Steuer spart jeder zusätzliche Euro Werbungskosten (bei kleinen Beträgen).")
    C("Durchschnittssatz", "Durchschnittssteuersatz", "IF({zvE}>0,{Steuer_zvE}/{zvE},0)", F_PCT)
    C("Steuer_Faktor_direkt", "Steuerfaktor im Modus 'direkt'", f'{{Grenzsteuersatz_direkt}}*(1+IF({{Soli_direkt}}="{JA}",{{P_Soli_Satz}},0)+{{KiSt_Satz_eff}})', F_PCT)
    C("Satz_verwendet", "Verwendeter Grenzsteuersatz (Info)", f'IF({{Steuer_Modus}}="{TARIF}",{{Grenzsteuersatz_calc}},{{Steuer_Faktor_direkt}})', F_PCT, key=True)

    r += 1
    m.section(ws, r, "Vermietungsanteil & Abzugsquoten"); r += 1
    C("Verm_Anteil", "Vermieteter Anteil", f'IF({{Modus}}="{KA}",1,{{Teilverm_Anteil}})', F_PCT, "Kapitalanlage = 100 %, Eigennutzung = Teilvermietungsanteil.")
    C("Miete_Start", "Kaltmiete Start (vermieteter Teil)", f'IF({{Modus}}="{KA}",{{Kaltmiete}},{{Teilverm_Miete}})', F_EUR, unit="€/Monat")
    C("Zins_Abzugsquote", "Abziehbarer Anteil der Schuldzinsen", f'IF({{Modus}}="{KA}",1,IF(AND({{Teilverm_Zuordnung}}="{JA}",{{Darlehen_Summe}}>0),MIN(1,{{Verm_Anteil}}*{{Gesamtkosten}}/{{Darlehen_Summe}}),{{Verm_Anteil}}))', F_PCT,
      "Bei gezielter Darlehenszuordnung zum vermieteten Teil höher als der Flächenanteil.")
    C("Miete_Quote", "Miete in % der ortsüblichen Miete", f'IF(AND({{Angehoerige}}="{JA}",{{Vergleichsmiete}}>0),{{Miete_Start}}/{{Vergleichsmiete}},1)', F_PCT)
    C("TUEP", "Totalüberschussprognose (30 J., ohne Sonder-AfA)", "SUM([[TUEP_Beitrag]])", F_EUR, "Nur relevant bei 50–66 % der ortsüblichen Miete.")
    C("WK_Quote", "Abziehbarer Anteil der Werbungskosten", f'IF({{Angehoerige}}<>"{JA}",1,IF({{Miete_Quote}}>={{P_Verb_Voll}},1,IF({{Miete_Quote}}>={{P_Verb_Min}},IF({{TUEP}}>0,1,{{Miete_Quote}}),{{Miete_Quote}})))', F_PCT, key=True)
    C("W_Verbilligt", "Prüfung verbilligte Vermietung",
      f'IF({{Angehoerige}}<>"{JA}","– nicht relevant",IF({{Miete_Quote}}>={{P_Verb_Voll}},"✔ ≥ 66 %: voller Werbungskostenabzug",IF({{Miete_Quote}}>={{P_Verb_Min}},IF({{TUEP}}>0,"⚠ 50–66 %: Totalüberschussprognose positiv → voller Abzug (dokumentieren!)","⚠ 50–66 %: Prognose negativ → Werbungskosten nur anteilig"),"✖ < 50 %: Werbungskosten nur anteilig abziehbar")))', None)

    r += 1
    m.section(ws, r, "AfA-Ermittlung (Gebäude)"); r += 1
    C("AfA_Satz_Std", "AfA-Satz nach Fertigstellungsjahr", "IF({Baujahr}<{P_AfA_Grenze1},{P_AfA_vor1925},IF({Baujahr}<{P_AfA_Grenze2},{P_AfA_1925_2022},{P_AfA_ab2023}))", F_PCT2)
    C("RND_aktiv", "RND-Gutachten wird angesetzt (1 = ja)", f'IF(AND({{RND_Gutachten}}="{JA}",{{RND_Jahre}}>0),IF(1/{{RND_Jahre}}>{{AfA_Satz_Std}},1,0),0)', F_INT,
      "Nur wenn die Gutachten-AfA höher ist als die Regel-AfA.")
    C("AfA_Satz_linear", "Linearer AfA-Satz (verwendet)", "IF({RND_aktiv}=1,1/{RND_Jahre},{AfA_Satz_Std})", F_PCT2, key=True)
    C("ND_linear", "Nutzungsdauer (Jahre)", "1/{AfA_Satz_linear}", F_NUM1)
    C("Degr_ber", "Degressive AfA zulässig (1 = ja)", f'IF(AND({{Baubeginn}}>={{P_Degr_Beginn}},{{Baubeginn}}<={{P_Degr_Ende}},{{Erwerb_FJ}}="{JA}"),1,0)', F_INT)
    C("Degr_aktiv", "Degressive AfA aktiv (1 = ja)", f'IF(AND({{Degr_Wahl}}="{JA}",{{Degr_ber}}=1,{{RND_aktiv}}=0),1,0)', F_INT)
    C("BK_m2", "Anschaffungskosten Gebäude je m²", "{AK_Gebaeude}/{Wohnflaeche}", F_EUR, unit="€/m²")
    C("Sonder7b_ber", "§7b zulässig (1 = ja)", f'IF(AND({{Bauantrag}}>={{P_7b_Antrag_Beginn}},{{Bauantrag}}<={{P_7b_Antrag_Ende}},{{Effizienz_EH40}}="{JA}",{{BK_m2}}<={{P_7b_BKO}},{{Erwerb_FJ}}="{JA}"),1,0)', F_INT)
    C("Sonder7b_aktiv", "§7b aktiv (1 = ja)", f'IF(AND({{Sonder7b_Wahl}}="{JA}",{{Sonder7b_ber}}=1),1,0)', F_INT)
    C("Sonder7b_BMG", "§7b Bemessungsgrundlage", "MIN({AK_Gebaeude},{P_7b_BMG_max}*{Wohnflaeche})", F_EUR)
    C("Denkmal_BMG", "Bemessungsgrundlage §7i/§7h/§10f", f'IF({{Denkmal_Status}}<>"{DENK_NEIN}",MIN({{Denkmal_Betrag}},{{AK_Gebaeude}}),0)', F_EUR)
    C("AN_Summe", "Instandsetzung in den ersten 3 Jahren (netto)", f'(IF({{Renov_Art}}="{ERH}",{{Renovierung}},0)+IF({{Erh_Jahr}}<={{P_AN_Jahre}},{{Erh_Betrag}},0))/(1+{{P_USt}})', F_EUR)
    C("AN_Grenze", "15-%-Grenze anschaffungsnaher Aufwand", "{P_AN_Quote}*{AK_Gebaeude}", F_EUR)
    C("AN_ueber", "15-%-Grenze überschritten (1 = ja)", "IF(AND({Verm_Anteil}>0,{AN_Summe}>{AN_Grenze}),1,0)", F_INT, key=True)
    C("Renov_HK", "Start-Renovierung als Herstellungskosten", f'IF(OR({{Renov_Art}}="{HK}",AND({{AN_ueber}}=1,{{Renov_Art}}="{ERH}")),{{Renovierung}},0)', F_EUR)
    C("Renov_Erhaltung", "Start-Renovierung als Erhaltungsaufwand", "{Renovierung}-{Renov_HK}", F_EUR)
    C("Erh_gross_HK", "Geplanter Erhaltungsaufwand → HK (anschaffungsnah)", "IF(AND({AN_ueber}=1,{Erh_Jahr}<={P_AN_Jahre}),{Erh_Betrag},0)", F_EUR)
    C("AfA_BMG_Gebaeude", "AfA-Bemessungsgrundlage Gebäude (Start)", "{AK_Gebaeude}-{Denkmal_BMG}+{Renov_HK}", F_EUR, key=True)
    C("Verteilung_eff", "Verteilung Erhaltungsaufwand (Jahre)", "IF({Verm_Anteil}>0,MAX(1,MIN({P_82b_max},{Verteilung_82b})),1)", F_INT)
    C("Ruecklage_sofort_f", "Rücklage sofort abziehbar (1 = ja)", f'IF({{Ruecklage_sofort}}="{JA}",1,0)', F_INT)
    C("AfA_Art", "Angewandte Gebäude-AfA",
      'IF({Degr_aktiv}=1,"Degressiv 5 % (§7 Abs. 5a)",IF({RND_aktiv}=1,"Linear nach Gutachten-RND",FIXED({AfA_Satz_Std}*100,1)&" %"&" linear (§7 Abs. 4)"))&IF({Sonder7b_aktiv}=1," + Sonder-AfA §7b","")&IF({Denkmal_BMG}>0," + §7i/§7h","")', None, key=True)

    r += 1
    m.section(ws, r, "Eigennutzung: Steuerermäßigungen"); r += 1
    C("Bonus35c_aktiv", "§35c anwendbar (1 = ja)", f'IF(AND({{Modus}}="{EN}",{{Alter}}>{{P_35c_Alter}},{{Energ_gefoerdert}}="{NEIN}",{{Startjahr}}<={{P_35c_Ende}},{{Renov_Energ}}>0),1,0)', F_INT)
    C("Bonus35c_BMG", "§35c begünstigte Aufwendungen (selbst genutzt)", "MIN({Renovierung}*{Renov_Energ},{P_35c_MaxAufw})*(1-{Verm_Anteil})", F_EUR)
    C("Renov_35a_Lohn", "§35a: Lohnanteil Start-Renovierung (selbst genutzt)",
      f'IF({{Modus}}="{EN}",{{Renovierung}}*(1-IF({{Bonus35c_aktiv}}=1,{{Renov_Energ}},0))*{{Renov_Lohn}}*(1-{{Verm_Anteil}}),0)', F_EUR)

    m.st_row = r + 1
    return ws


def build_szen_scalars(m):
    ws = m.sheet("sz", "Szenarien & Sensitivität")
    m.title(ws, "Szenarien & Sensitivität",
            "Best/Base/Worst: Abweichungen ggü. deinen Eingaben. Aktives Szenario wählst du auf 'Eingaben'. "
            "Die Raster rechnen das komplette Modell für jede Zelle neu (ausgeblendetes Blatt _Sens).")
    for col, w in zip("BCDEFGH", (46, 22, 16, 16, 16, 16, 50)):
        ws.column_dimensions[col].width = w
    r = 4
    m.section(ws, r, "Szenario-Definition (Base = deine Eingaben)", ncols=8); r += 1
    m.subhead(ws, r, ["Stellschraube", "Base", "Best", "Worst", "", "", "Erklärung"]); r += 1
    rows = [
        ("SzT_Anschluss", "Anschlusszins (Δ Prozentpunkte)", (0, -0.01, 0.02), F_PCT2, ("dec", -0.1, 0.1, "Bitte −10 % bis +10 % eingeben."), "Verschiebt Anschlusszinsen und Nachbeleihungszins."),
        ("SzT_Soll", "Sollzins (Δ, falls noch nicht fixiert)", (0, 0, 0), F_PCT2, ("dec", -0.1, 0.1, "Bitte −10 % bis +10 % eingeben."), "Nur ändern, wenn der Zins noch nicht vertraglich fest ist."),
        ("SzT_Miete", "Startmiete (Faktor)", (1, 1.0, 0.95), F_PCT, ("dec", 0, 3, "Bitte 0 % bis 300 % eingeben."), "Auch für die Vergleichsmiete beim Mieten."),
        ("SzT_Mietsteig", "Mietsteigerung (Δ p.a.)", (0, 0.005, -0.01), F_PCT2, ("dec", -0.1, 0.1, "Bitte −10 % bis +10 % eingeben."), ""),
        ("SzT_Leerstand", "Leerstand (Δ Prozentpunkte)", (0, -0.01, 0.05), F_PCT2, ("dec", -1, 1, "Bitte −100 % bis +100 % eingeben."), ""),
        ("SzT_Wert", "Wertsteigerung (Δ p.a.)", (0, 0.01, -0.015), F_PCT2, ("dec", -0.1, 0.1, "Bitte −10 % bis +10 % eingeben."), ""),
        ("SzT_Instand", "Instandhaltung (Faktor)", (1, 0.8, 1.5), F_PCT, ("dec", 0, 5, "Bitte 0 % bis 500 % eingeben."), "Skaliert die eigene Instandhaltungsreserve."),
        ("SzT_Inflation", "Kosteninflation (Δ p.a.)", (0, 0, 0.01), F_PCT2, ("dec", -0.1, 0.1, "Bitte −10 % bis +10 % eingeben."), ""),
    ]
    for name, label, vals, fmt, spec, hint in rows:
        m.label(ws, r, label, None, hint, hint_col=8)
        c = ws.cell(row=r, column=3, value=vals[0])
        m.style(c, fmt=fmt, fillc=C_GREY)
        for i in (1, 2):
            m.inp(ws, r, f"{name}_{'Best' if i == 1 else 'Worst'}", label, vals[i], fmt, spec, None,
                  col=3 + i, write_label=False)
        m.define(name, ws, f"$C${r}:$E${r}")
        m.leaf(name)
        r += 1
    r += 1
    m.section(ws, r, "Wirksame Werte im aktiven Szenario", ncols=8); r += 1

    def C(name, label, tpl, fmt, hint=None):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, None)
        if hint:
            m.text(ws, r, 8, hint, size=9, color="404040", border=True)
        r += 1

    C("Szen_Idx", "Szenario-Index (1 Base, 2 Best, 3 Worst)", "MATCH({Szenario},{Liste_Szenario},0)", F_INT)
    C("Szen_Anschluss_Delta", "Δ Anschlusszins", "INDEX({SzT_Anschluss},{Szen_Idx})", F_PCT2)
    C("Szen_Soll_Delta", "Δ Sollzins", "INDEX({SzT_Soll},{Szen_Idx})", F_PCT2)
    C("Szen_Miete_Faktor", "Faktor Startmiete", "INDEX({SzT_Miete},{Szen_Idx})", F_PCT)
    C("Szen_Instand_Faktor", "Faktor Instandhaltung", "INDEX({SzT_Instand},{Szen_Idx})", F_PCT)
    C("Sens_Zins_Delta", "Zinsverschiebung Sensitivität (Basis 0)", "0", F_PCT2, "Wird nur in der Schattenrechnung verändert.")
    C("Miete_Start_eff", "Startmiete wirksam (vermieteter Teil)", "{Miete_Start}*{Szen_Miete_Faktor}", F_EUR)
    C("Miete_Alt_Start", "Vergleichsmiete Mieten wirksam (p.a.)", f'IF({{Modus}}="{EN}",{{Miete_Alternativ}}*12*{{Szen_Miete_Faktor}},0)', F_EUR)
    C("Mietsteig_eff", "Mietsteigerung wirksam", "{Mietsteigerung}+INDEX({SzT_Mietsteig},{Szen_Idx})", F_PCT2)
    C("Leerstand_eff", "Leerstand wirksam", "MAX(0,MIN(1,{Leerstand}+INDEX({SzT_Leerstand},{Szen_Idx})))", F_PCT2)
    C("Wertsteig_eff", "Wertsteigerung wirksam", "{Wertsteigerung}+INDEX({SzT_Wert},{Szen_Idx})", F_PCT2)
    C("Inflation_eff", "Kosteninflation wirksam", "{Inflation}+INDEX({SzT_Inflation},{Szen_Idx})", F_PCT2)
    m.sz_row = r + 1
    return ws


def build_kosten_scalars(m):
    """Laufende Startwerte (auf Blatt Cashflow-Plan oberhalb der Tabelle wäre unübersichtlich -> Steuer-/Fin-Blatt)."""
    ws = m.sheets["fin"]
    # wird unten am Finanzierungsblatt angehängt
    r = ws.max_row + 2
    m.section(ws, r, "Laufende Kosten im Jahr 1 (Basis für Inflation)"); r += 1

    def C(name, label, tpl, fmt, hint=None, unit=None):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, hint, unit=unit)
        r += 1

    C("U_Start", "Umlagefähige Kosten p.a.", "{HG_umlage}*12+{Grundsteuer}+{Vers_umlage}", F_EUR)
    C("N_Start", "Nicht umlagefähige Kosten p.a. (ohne Rücklage)", "{HG_nicht_umlage}*12+{Vers_nicht_umlage}+{Mietverwaltung}*12", F_EUR)
    C("R_Start", "Zuführung Erhaltungsrücklage p.a.", "{HG_Ruecklage}*12", F_EUR)
    C("Peters_m2", "Vorschlag Instandhaltung (Peters) je m²", f'{{Peters_HK}}*{{P_Peters_Faktor}}/{{P_Peters_Jahre}}*IF({{Objekttyp}}="{ETW}",{{P_Peters_ETW}},1)*INDEX({{Tab_Zust_Inst}},MATCH({{Zustand}},{{Liste_Zustand}},0))', F_EUR2, unit="€/m²·a")
    C("Instand_pa", "Eigene Instandhaltungsreserve p.a.", f'IF({{Inst_Modus}}="{INST_P}",{{Peters_m2}},{{Inst_manuell}})*{{Wohnflaeche}}', F_EUR)
    C("S_Start", "Sonstige Werbungskosten p.a.", "{Fahrt_km}*{P_km_Satz}+{Kontofuehrung}+{Steuerberatung}+{Sonstige_WK}", F_EUR)
    C("NB_an", "Nachbeleihung aktiv (1 = ja)", f'IF({{NB_Aktiv}}="{JA}",1,0)', F_INT)
    C("NB_Mod_HK", "Nachbeleihung-Modernisierung = Herstellung (1 = ja)", f'IF({{NB_Mod_Art}}="{HK}",1,0)', F_INT)
    C("NB_Abzug_Quote", "Abziehbarer Anteil Zinsen Nachbeleihung",
      f'IF({{NB_Verwendung}}="{NB_MOD}",{{Zins_Abzugsquote}},IF({{NB_Verwendung}}="{NB_FOLGE}",1,0))', F_PCT,
      "Verwendungszweck entscheidet (Veranlassungsprinzip).")


# ====================================================================== Jahresmodell
def engine_columns():
    E = []

    def col(key, header, tpl, init=None, fmt=F_EUR, group="", width=12, const=False):
        E.append(dict(key=key, header=header, tpl=tpl, init=init, fmt=fmt, group=group,
                      width=width, const=const))

    g = "Übersicht"
    col("t", "Jahr", None, fmt=F_INT, group=g, width=6, const=True)
    col("Kal", "Kalender-jahr", "{Startjahr}+[t]-1", fmt=F_INT, group=g, width=8)
    col("Aktiv", "im Horizont", "IF([t]<={Horizont},1,0)", fmt=F_INT, group=g, width=7)
    col("O_Miete", "Miete (netto Ausfall)", "[Miete_Ist]", group=g)
    col("O_Bewirt", "Bewirtschaftung & Instandhaltung", "-([U_Traeger]+[N_Kosten]+[Ruecklage]+[Instand]+[Erh_gross]+[S_Kosten]+[Bereit])", group=g)
    col("O_Zins", "Zinsen", "-[Zinsen_Summe]", group=g)
    col("O_Tilg", "Tilgung (inkl. Sonder)", "-[Tilgung_Summe]", group=g)
    col("CF_vor", "Cashflow vor Steuern",
        "[Miete_Ist]-[U_Traeger]-[N_Kosten]-[Ruecklage]-[Instand]-[Erh_gross]-[S_Kosten]-[Bereit]-[Kapitaldienst]", group=g)
    col("O_Steuer", "Steuer (+ Erstattung / − Last)", "-[Steuerwirkung]", group=g)
    col("CF_nach", "Cashflow nach Steuern", "[CF_vor]-[Steuerwirkung]", group=g)
    col("CF_mtl", "… pro Monat", "[CF_nach]/12", group=g, width=10)
    col("CF_kum", "Cashflow kumuliert", "[CF_kum@-1]+[CF_nach]", init="0", group=g)
    col("Wert", "Immobilien-wert", "[Wert@-1]*(1+{Wertsteig_eff})+[NB_Mod]*{NB_Wertanteil}", init="{Startwert}", group=g)
    col("RS_Summe", "Restschuld gesamt", "[L1_RS_E]+[L2_RS_E]+[L3_RS_E]+[L4_RS_E]+[L5_RS_E]",
        init="[L1_RS_E]+[L2_RS_E]+[L3_RS_E]+[L4_RS_E]+[L5_RS_E]", group=g)
    col("Vermoegen", "Vermögen (Wert − Schuld + kum. CF)", "[Wert]-[RS_Summe]+[CF_kum]", init="[Wert]-[RS_Summe]", group=g)

    g = "Einnahmen & Kosten"
    col("Infl", "Kosten-index", "(1+{Inflation_eff})^([t]-1)", fmt=F_NUM2, group=g, width=8)
    col("Miete_Soll", "Sollmiete p.a.", "{Miete_Start_eff}*12*(1+{Mietsteig_eff})^([t]-1)", group=g)
    col("Mietausfall", "Mietausfall", "[Miete_Soll]*{Leerstand_eff}", group=g)
    col("Miete_Ist", "Mieteinnahmen", "[Miete_Soll]-[Mietausfall]", group=g)
    col("U_Kosten", "Umlagefähige Kosten", "{U_Start}*[Infl]", group=g)
    col("U_Traeger", "davon selbst getragen", "[U_Kosten]*({Verm_Anteil}*{Leerstand_eff}+1-{Verm_Anteil})", group=g)
    col("N_Kosten", "Nicht umlagefähig", "{N_Start}*[Infl]", group=g)
    col("Ruecklage", "Rücklage WEG", "{R_Start}*[Infl]", group=g)
    col("Instand", "Instandhaltung eigen", "{Instand_pa}*{Szen_Instand_Faktor}*[Infl]", group=g)
    col("Erh_gross", "Größerer Erhaltungs-aufwand", "IF([t]={Erh_Jahr},{Erh_Betrag},0)", group=g)
    col("S_Kosten", "Sonstige WK (Fahrt, Konto, StB …)", "IF({Verm_Anteil}>0,{S_Start}*[Infl],0)", group=g)
    col("Bereit", "Bereitstellungs-zinsen", "IF([t]=1,{Bereitstellung},0)", group=g)

    for k in range(1, 5):
        g = f"Darlehen {k}"
        K = f"Kredit{k}"
        L = f"L{k}"
        col(f"{L}_RS_A", "RS Anfang", f"[{L}_RS_E@-1]", group=g)
        col(f"{L}_Zins", "Zinssatz",
            f"MAX(0,IF([t]<={{{K}_ZB}},{{{K}_Zins}}+{{Szen_Soll_Delta}},{{{K}_Anschluss}}+{{Szen_Anschluss_Delta}})+{{Sens_Zins_Delta}})",
            fmt=F_PCT2, group=g, width=8)
        col(f"{L}_Ann", "Jahresrate (Soll)",
            f"IF([{L}_RS_A]<=0,0,IF([t]<={{{K}_TF}},[{L}_RS_A]*[{L}_Zins],IF([t]<={{{K}_ZB}},{{{K}_Betrag}}*([{L}_Zins]+{{{K}_Tilgung}}),"
            f"IF([t]={{{K}_ZB}}+1,[{L}_RS_A]*([{L}_Zins]+{{{K}_AT}}),[{L}_Ann@-1]))))", init="0", group=g)
        col(f"{L}_RS_reg", "RS nach Raten",
            f"IF([{L}_Zins]=0,MAX(0,[{L}_RS_A]-[{L}_Ann]),MAX(0,[{L}_RS_A]*(1+[{L}_Zins]/12)^12-[{L}_Ann]/12*((1+[{L}_Zins]/12)^12-1)/([{L}_Zins]/12)))",
            group=g)
        col(f"{L}_Tilg", "Tilgung", f"[{L}_RS_A]-[{L}_RS_reg]", group=g)
        col(f"{L}_Zinsen", "Zinsen", f"MAX(0,MIN([{L}_Ann]-[{L}_Tilg],[{L}_RS_A]*[{L}_Zins]))", group=g)
        col(f"{L}_Sonder", "Sonder-tilgung",
            f"IF(AND([t]>={{{K}_ST_ab}},[t]<={{{K}_ST_bis}}),MIN({{{K}_ST}},[{L}_RS_reg]),0)", group=g)
        col(f"{L}_Zuschuss", "Tilgungs-zuschuss",
            f"IF([t]={{{K}_TZ_Jahr}},MAX(0,MIN({{{K}_TZ}}*{{{K}_Betrag}},[{L}_RS_reg]-[{L}_Sonder])),0)", group=g)
        col(f"{L}_RS_E", "RS Ende", f"[{L}_RS_reg]-[{L}_Sonder]-[{L}_Zuschuss]", init=f"{{{K}_Betrag}}", group=g)

    g = "Nachbeleihung"
    col("NB_Ausz", "Auszahlung",
        f'IF(AND({{NB_an}}=1,[t]={{NB_Jahr}}),IF({{NB_Modus}}="{NB_FIX}",{{NB_Betrag}},MAX(0,{{NB_Ziel}}*[Wert@-1]-[RS_Summe@-1])),0)', group=g)
    col("L5_RS_A", "RS Anfang", "[L5_RS_E@-1]+[NB_Ausz]", group=g)
    col("L5_Zins", "Zinssatz", "MAX(0,{NB_Zins}+{Szen_Anschluss_Delta}+{Sens_Zins_Delta})", fmt=F_PCT2, group=g, width=8)
    col("L5_Ann", "Jahresrate", "IF([L5_RS_A]<=0,0,IF([NB_Ausz]>0,[L5_RS_A]*([L5_Zins]+{NB_Tilgung}),[L5_Ann@-1]))", init="0", group=g)
    col("L5_RS_E", "RS Ende",
        "IF([L5_Zins]=0,MAX(0,[L5_RS_A]-[L5_Ann]),MAX(0,[L5_RS_A]*(1+[L5_Zins]/12)^12-[L5_Ann]/12*((1+[L5_Zins]/12)^12-1)/([L5_Zins]/12)))",
        init="0", group=g)
    col("L5_Tilg", "Tilgung", "[L5_RS_A]-[L5_RS_E]", group=g)
    col("L5_Zinsen", "Zinsen", "MAX(0,MIN([L5_Ann]-[L5_Tilg],[L5_RS_A]*[L5_Zins]))", group=g)
    col("NB_Mod", "→ Modernisierung", f'IF({{NB_Verwendung}}="{NB_MOD}",[NB_Ausz],0)', group=g)
    col("NB_Entn", "→ entnommen", "[NB_Ausz]-[NB_Mod]", group=g)
    col("NB_Konsum", "→ Konsum", f'IF({{NB_Verwendung}}="{NB_KONSUM}",[NB_Ausz],0)', group=g)
    col("NB_Konsum_kum", "Konsum kumuliert", "[NB_Konsum_kum@-1]+[NB_Konsum]", init="0", group=g)
    col("NB_Topf", "Wert Folgeobjekt-EK / ETF",
        f'[NB_Topf@-1]*(1+IF({{NB_Verwendung}}="{NB_ETF}",{{ETF_Rendite}},{{NB_Folge_Rendite}}))+[NB_Entn]-[NB_Konsum]', init="0", group=g)
    col("NB_Topf_Basis", "Einzahlungen", "[NB_Topf_Basis@-1]+[NB_Entn]-[NB_Konsum]", init="0", group=g)
    col("NB_HK_kum", "Modernisierung als HK kum.", "[NB_HK_kum@-1]+IF({NB_Mod_HK}=1,[NB_Mod],0)", init="0", group=g)

    g = "Darlehen gesamt"
    col("Zinsen_Summe", "Zinsen gesamt", "[L1_Zinsen]+[L2_Zinsen]+[L3_Zinsen]+[L4_Zinsen]+[L5_Zinsen]", group=g)
    col("Tilgung_Summe", "Tilgung gesamt",
        "[L1_Tilg]+[L1_Sonder]+[L2_Tilg]+[L2_Sonder]+[L3_Tilg]+[L3_Sonder]+[L4_Tilg]+[L4_Sonder]+[L5_Tilg]", group=g)
    col("Zuschuss_Summe", "Tilgungs-zuschüsse", "[L1_Zuschuss]+[L2_Zuschuss]+[L3_Zuschuss]+[L4_Zuschuss]", group=g)
    col("Kapitaldienst", "Kapitaldienst", "[Zinsen_Summe]+[Tilgung_Summe]", group=g)
    dis = "+".join(
        f"IF(AND({{Kredit{k}_Disagio}}<={{P_Disagio_max}},{{Kredit{k}_ZB}}>={{P_Disagio_ZB_min}}),IF([t]=1,{{Kredit{k}_Betrag}}*{{Kredit{k}_Disagio}},0),"
        f"IF([t]<={{Kredit{k}_ZB}},{{Kredit{k}_Betrag}}*{{Kredit{k}_Disagio}}/{{Kredit{k}_ZB}},0))" for k in range(1, 5))
    col("Disagio_Abzug", "Disagio (WK)", dis, group=g)

    g = "Abschreibung"
    col("Geb_BMG", "AfA-Basis Gebäude", "{AfA_BMG_Gebaeude}+IF([t]>={Erh_Jahr},{Erh_gross_HK},0)+[NB_HK_kum]", group=g)
    col("Geb_RW", "Restwert vor AfA", "[Geb_BMG]-[Geb_AfA_kum@-1]-[S7b_kum@-1]", group=g)
    col("Geb_AfA", "Gebäude-AfA",
        "MAX(0,MIN([Geb_RW],IF({Degr_aktiv}=1,MAX({P_Degr_Satz}*[Geb_RW],[Geb_RW]/MAX(1,{ND_linear}-[t]+1)),"
        "IF(AND({Sonder7b_aktiv}=1,[t]>{P_7b_Jahre}),[Geb_RW]/MAX(1,{ND_linear}-[t]+1),{AfA_Satz_linear}*[Geb_BMG]))))", group=g)
    col("S7b", "Sonder-AfA §7b",
        "IF(AND({Sonder7b_aktiv}=1,[t]<={P_7b_Jahre}),MAX(0,MIN({P_7b_Satz}*{Sonder7b_BMG},[Geb_RW]-[Geb_AfA])),0)", group=g)
    col("Geb_AfA_kum", "AfA kumuliert", "[Geb_AfA_kum@-1]+[Geb_AfA]", init="0", group=g)
    col("S7b_kum", "§7b kumuliert", "[S7b_kum@-1]+[S7b]", init="0", group=g)
    col("Denkmal_AfA", "§7i/§7h-AfA",
        "IF([t]<={P_7i_Jahre1},{P_7i_Satz1},IF([t]<={P_7i_Jahre1}+{P_7i_Jahre2},{P_7i_Satz2},0))*{Denkmal_BMG}", group=g)
    col("S10f", "§10f Sonder-ausgabe",
        f'IF(AND({{Modus}}="{EN}",[t]<={{P_10f_Jahre}}),{{Denkmal_BMG}}*{{P_10f_Satz}}*(1-{{Verm_Anteil}}),0)', group=g)
    col("Inventar_AfA", "Inventar-AfA", "IF([t]<={Inventar_ND},{Inventar}/{Inventar_ND},0)", group=g)
    col("AfA_voll", "AfA gesamt (100 %)", "[Geb_AfA]+[S7b]+[Denkmal_AfA]+[Inventar_AfA]", group=g)
    col("AfA_abg_kum", "Gebäude-AfA abgezogen kum.",
        "[AfA_abg_kum@-1]+([Geb_AfA]+[S7b]+[Denkmal_AfA])*{Verm_Anteil}*{WK_Quote}", init="0", group=g)

    g = "Werbungskosten & Einkünfte V+V"
    col("Renov_Abzug", "Start-Renovierung (Erhaltung)", "IF([t]<={Verteilung_eff},{Renov_Erhaltung}/{Verteilung_eff},0)", group=g)
    col("Erh_Abzug", "Größerer Erhaltungs-aufwand",
        "IF(AND({Erh_gross_HK}=0,[t]>={Erh_Jahr},[t]<{Erh_Jahr}+{Verteilung_eff}),{Erh_Betrag}/{Verteilung_eff},0)", group=g)
    col("NB_Erh_Abzug", "Modernisierung NB (Erhaltung)",
        "IF(AND({NB_Mod_HK}=0,[t]>={NB_Jahr},[t]<{NB_Jahr}+{Verteilung_eff}),INDEX([[NB_Mod]],{NB_Jahr}+1)/{Verteilung_eff},0)", group=g)
    col("WK_Bewirt", "WK Bewirtschaftung",
        "{Verm_Anteil}*([U_Kosten]*{Leerstand_eff}+[N_Kosten]+[Ruecklage]*{Ruecklage_sofort_f}+[Instand]+[Renov_Abzug]+[Erh_Abzug]+[NB_Erh_Abzug])+[S_Kosten]", group=g)
    col("WK_Zinsen", "WK Schuldzinsen",
        "([L1_Zinsen]+[L2_Zinsen]+[L3_Zinsen]+[L4_Zinsen])*{Zins_Abzugsquote}+[L5_Zinsen]*{NB_Abzug_Quote}", group=g)
    col("WK_Finanz", "WK Finanzierungs-nebenkosten",
        "IF([t]=1,({NK_Finanzierung}+{Bereitstellung})*{Zins_Abzugsquote}+{RND_Kosten_eff}*{Verm_Anteil},0)+[Disagio_Abzug]*{Zins_Abzugsquote}", group=g)
    col("WK_AfA", "WK AfA", "[AfA_voll]*{Verm_Anteil}", group=g)
    col("WK_voll", "Werbungskosten gesamt", "[WK_Bewirt]+[WK_Zinsen]+[WK_Finanz]+[WK_AfA]", group=g)
    col("WK_abz", "davon abziehbar", "[WK_voll]*{WK_Quote}", group=g)
    col("Einnahmen", "Einnahmen V+V", "[Miete_Ist]+[Zuschuss_Summe]*{Verm_Anteil}", group=g)
    col("TUEP_Beitrag", "Beitrag Total-überschuss", "IF([t]<={P_TUEP_Jahre},[Einnahmen]-[WK_voll]+[S7b]*{Verm_Anteil},0)", group=g)
    col("VuV", "Einkünfte V+V", "[Einnahmen]-[WK_abz]", group=g)

    g = "Einkommensteuer"
    col("zvE_ohne", "zvE ohne Objekt", "{zvE}*(1+{zvE_Steigerung})^([t]-1)", group=g)
    col("ESt_ohne", "ESt ohne", T_tpl("[zvE_ohne]"), group=g)
    col("Steuer_ohne", "ESt+Soli+KiSt ohne", total_tax_tpl("[ESt_ohne]"), group=g)
    col("Basis_mit", "Einkünfte mit Objekt", "[zvE_ohne]+[VuV]-[S10f]", group=g)
    col("VV", "Verlust-vortrag", "MAX(0,[VV@-1]-MAX(0,[Basis_mit]))+MAX(0,-[Basis_mit])", init="0", group=g)
    col("zvE_mit", "zvE mit Objekt", "MAX(0,[Basis_mit]-[VV@-1])", group=g)
    col("ESt_mit_T", "ESt tariflich mit", T_tpl("[zvE_mit]"), group=g)
    col("Erm_35a", "§35a", f'IF({{Modus}}="{EN}",MIN({{P_35a_HW_Satz}}*({{Handwerker_Lohn}}*(1-{{Verm_Anteil}})+IF([t]=1,{{Renov_35a_Lohn}},0)),{{P_35a_HW_Max}})'
        f'+MIN({{P_35a_HHDL_Satz}}*{{HHDL}}*(1-{{Verm_Anteil}}),{{P_35a_HHDL_Max}}),0)', group=g)
    col("Erm_35c", "§35c", "IF({Bonus35c_aktiv}=1,{Bonus35c_BMG}*IF([t]=1,{P_35c_S1},IF([t]=2,{P_35c_S2},IF([t]=3,{P_35c_S3},0))),0)", group=g)
    col("ESt_mit", "ESt festgesetzt mit", "MAX(0,[ESt_mit_T]-[Erm_35a]-[Erm_35c])", group=g)
    col("Steuer_mit", "ESt+Soli+KiSt mit", total_tax_tpl("[ESt_mit]"), group=g)
    col("Steuerwirkung", "Mehr-/Minder-steuer (+ = Last)",
        f'IF({{Steuer_Modus}}="{TARIF}",[Steuer_mit]-[Steuer_ohne],([VuV]-[S10f])*{{Steuer_Faktor_direkt}}-([Erm_35a]+[Erm_35c])*(1+{{KiSt_Satz_eff}}))', group=g)

    g = "Vergleich & Rendite"
    col("Alt_Ausgabe", "Ausgaben Alternative (Miete)",
        f'IF({{Modus}}="{EN}",{{Miete_Alt_Start}}*(1+{{Mietsteig_eff}})^([t]-1)+[U_Kosten]*(1-{{Verm_Anteil}}),0)', group=g)
    col("Diff", "Mehrkosten Immobilie", "-[CF_nach]-[Alt_Ausgabe]", group=g)
    col("Depot_A", "Depot Käufer (Überschüsse)", "[Depot_A@-1]*(1+{ETF_Rendite})+MAX(0,-[Diff])", init="0", group=g)
    col("Basis_A", "Einzahlungen Käufer", "[Basis_A@-1]+MAX(0,-[Diff])", init="0", group=g)
    col("Depot_B", "Depot Alternative", "[Depot_B@-1]*(1+{ETF_Rendite})+MAX(0,[Diff])", init="MAX(0,{EK_Einsatz})", group=g)
    col("Basis_B", "Einzahlungen Alternative", "[Basis_B@-1]+MAX(0,[Diff])", init="MAX(0,{EK_Einsatz})", group=g)
    col("Verm_A", "Vermögen Immobilie (liquid., vor St.)",
        "[Wert]*(1-{Verkaufskosten})-[RS_Summe]+[Depot_A]+[NB_Topf]+[NB_Konsum_kum]",
        init="[Wert]*(1-{Verkaufskosten})-[RS_Summe]", group=g)
    col("Verm_B", "Vermögen Alternative (vor St.)", "[Depot_B]", init="[Depot_B]", group=g)
    col("Liq_Verm", "Liquidationswert inkl. kum. CF", "[Wert]*(1-{Verkaufskosten})-[RS_Summe]+[CF_kum]+[NB_Topf]+[NB_Konsum_kum]", group=g)
    col("IRR_Flow", "Zahlungsstrom IRR",
        "IF([t]<={Horizont},[CF_nach]+[Alt_Ausgabe]+[NB_Entn]+IF([t]={Horizont},{Verkauf_Netto},0),0)",
        init="-{EK_Einsatz}", group=g)
    col("BE_V", "Break-even Vermögen",
        f'IF(IF({{Modus}}="{KA}",[Liq_Verm]>={{EK_Einsatz}},[Verm_A]>=[Verm_B]),[t],999)', fmt=F_INT, group=g, width=8)
    col("BE_CF", "Break-even Cashflow", "IF([CF_nach]>=0,[t],999)", fmt=F_INT, group=g, width=8)
    return E


def build_cashflow(m):
    ws = m.sheet("cf", "Cashflow-Plan")
    m.title(ws, "Cashflow-Plan – Jahr für Jahr",
            "Komplettes Jahresmodell (bis 40 Jahre). Graue Zeilen liegen nach dem Betrachtungszeitraum. Spalten ab 'Einnahmen & Kosten' = Detailrechnung.")
    E = m.engine
    hdr_group, hdr, unit_row = 4, 5, 6
    row0 = 7
    m.eng_row0 = row0
    for i, c in enumerate(E):
        m.eng_col[c["key"]] = get_column_letter(i + 2)
    prev_group = None
    for i, c in enumerate(E):
        col = i + 2
        L = get_column_letter(col)
        ws.column_dimensions[L].width = c["width"]
        if c["group"] != prev_group:
            gcell = ws.cell(row=hdr_group, column=col, value=c["group"])
            gcell.font = fnt(bold=True, color="FFFFFF")
            prev_group = c["group"]
        ws.cell(row=hdr_group, column=col).fill = fill(C_HEAD)
        h = ws.cell(row=hdr, column=col, value=c["header"])
        m.style(h, bold=True, fillc=C_SUB, wrap=True, align="center", size=9)
        for t in range(0, T_MAX + 1):
            r = row0 + t
            if c["const"]:
                v = t
            elif t == 0:
                v = m.to_main(c["init"], r) if c["init"] is not None else None
            else:
                v = m.to_main(c["tpl"], r)
            cell = ws.cell(row=r, column=col, value=v)
            cell.number_format = c["fmt"]
            cell.font = fnt(size=9, bold=c["key"] in ("CF_vor", "CF_nach", "t"))
            cell.border = BORDER
            if c["key"] in ("CF_nach", "Vermoegen"):
                cell.fill = fill(C_KEY)
        m.define("Z_" + c["key"], ws, f"${L}${row0}:${L}${row0 + T_MAX}")
        m.nodes["Z_" + c["key"]] = {"tpl": None, "kind": "engine"}
    ws.row_dimensions[hdr].height = 45
    m.text(ws, row0 - 1, 1, "")
    for col in range(2, len(E) + 2):
        ws.cell(row=row0, column=col).fill = fill(C_GREY)
    ws.cell(row=row0 - 1, column=2, value="Jahr 0 = Kaufzeitpunkt").font = fnt(italic=True, size=8, color="7F7F7F")
    last = get_column_letter(len(E) + 1)
    ws.conditional_formatting.add(f"B{row0 + 1}:{last}{row0 + T_MAX}",
                                  FormulaRule(formula=[f"$B{row0 + 1}>Horizont"], font=Font(color="A6A6A6"),
                                              fill=PatternFill("solid", fgColor="FAFAFA")))
    ws.freeze_panes = ws.cell(row=row0, column=4)
    m.protect(ws)


# ====================================================================== Auswertung (Post-Skalare)
def build_steuer_post(m):
    ws = m.sheets["st"]
    r = m.st_row

    def C(name, label, tpl, fmt, hint=None, key=False, unit=None):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, hint, key=key, unit=unit)
        r += 1

    m.section(ws, r, "Verkauf am Ende des Betrachtungszeitraums (§23 EStG)"); r += 1
    C("Wert_N", "Verkaufspreis (Marktwert Ende Horizont)", "INDEX([[Wert]],{Horizont}+1)", F_EUR)
    C("VK_N", "Verkaufskosten", "{Wert_N}*{Verkaufskosten}", F_EUR)
    C("RS_N", "Restschuld bei Verkauf", "INDEX([[RS_Summe]],{Horizont}+1)", F_EUR)
    vfe = "+".join(f"IF({{Horizont}}<MIN({{Kredit{k}_ZB}},{{P_Kuend_489}}),INDEX([[L{k}_RS_E]],{{Horizont}}+1),0)" for k in range(1, 5))
    C("VFE_Betrag", "Vorfälligkeitsentschädigung (Schätzung)", f"{{VFE_Satz}}*({vfe})", F_EUR)
    C("Verkauf_Netto_vSt", "Netto-Erlös nach Schuldenablösung (vor Steuer)", "{Wert_N}-{VK_N}-{RS_N}-{VFE_Betrag}", F_EUR)
    C("AK_gesamt", "Anschaffungs-/Herstellungskosten gesamt", "{AK_Immobilie}+{Renov_HK}+IF({Erh_Jahr}<={Horizont},{Erh_gross_HK},0)+INDEX([[NB_HK_kum]],{Horizont}+1)", F_EUR)
    C("AfA_abg_N", "In Anspruch genommene Gebäude-AfA (Rückrechnung)", "INDEX([[AfA_abg_kum]],{Horizont}+1)", F_EUR,
      "Wird dem Veräußerungsgewinn hinzugerechnet (§23 Abs. 3 S. 4 EStG).")
    C("Gewinn23_roh", "Veräußerungsgewinn (vermieteter Teil)", "{Verm_Anteil}*({Wert_N}-{VK_N}-{AK_gesamt})+{AfA_abg_N}", F_EUR)
    C("Gewinn23_pflichtig", "Innerhalb der Spekulationsfrist (1 = ja)", "IF({Horizont}<{P_Spekfrist},1,0)", F_INT,
      "Eigennutzung: selbst genutzter Teil steuerfrei (Verkaufsjahr + 2 Vorjahre selbst genutzt).")
    C("Gewinn23", "Steuerpflichtiger Veräußerungsgewinn", "IF(AND({Gewinn23_pflichtig}=1,{Gewinn23_roh}>={P_23_Freigrenze}),{Gewinn23_roh},0)", F_EUR, key=True,
      hint="Gewinne unter 1.000 € (Freigrenze) bleiben steuerfrei. Verluste nur mit §23-Gewinnen verrechenbar (hier nicht angesetzt).")
    C("zvE_N", "zvE im Verkaufsjahr (ohne Gewinn)", "INDEX([[zvE_mit]],{Horizont}+1)", F_EUR)
    C("Erm_N", "Steuerermäßigungen im Verkaufsjahr", "INDEX([[Erm_35a]],{Horizont}+1)+INDEX([[Erm_35c]],{Horizont}+1)", F_EUR)
    C("EStV0", "ESt ohne Veräußerungsgewinn", f"MAX(0,{T_tpl('{zvE_N}')}-{{Erm_N}})", F_EUR)
    C("EStV1", "ESt mit Veräußerungsgewinn", f"MAX(0,{T_tpl('{zvE_N}+{Gewinn23}')}-{{Erm_N}})", F_EUR)
    C("Steuer_Verkauf", "Steuer auf den Veräußerungsgewinn",
      f'IF({{Steuer_Modus}}="{TARIF}",({total_tax_tpl("{EStV1}")})-({total_tax_tpl("{EStV0}")}),{{Gewinn23}}*{{Steuer_Faktor_direkt}})', F_EUR, key=True)
    C("Verkauf_Netto", "Netto-Erlös nach Steuern", "{Verkauf_Netto_vSt}-{Steuer_Verkauf}", F_EUR, key=True)
    C("W_DreiObj", "Drei-Objekt-Grenze",
      'IF({Horizont}>{P_DreiObj_Jahre},"✔ Haltedauer > 5 Jahre: zählt nicht zur Drei-Objekt-Grenze",IF({Objektverkaeufe}>{P_DreiObj},"✖ Mehr als 3 Objektverkäufe binnen 5 Jahren → gewerblicher Grundstückshandel droht (GewSt, keine Spekulationsfrist)!",IF({Objektverkaeufe}={P_DreiObj},"⚠ Genau 3 Objekte – kein weiteres Objekt binnen 5 Jahren verkaufen.","✔ Unkritisch")))', None)
    C("W_Spek", "Spekulationsfrist",
      f'IF({{Modus}}="{EN}",IF({{Verm_Anteil}}>0,IF({{Gewinn23_pflichtig}}=1,"⚠ Selbst genutzter Teil steuerfrei, vermieteter Teil steuerpflichtig (< 10 Jahre).","✔ Steuerfrei (> 10 Jahre)."),"✔ Eigennutzung: Verkauf steuerfrei (Verkaufsjahr + 2 Vorjahre selbst genutzt)."),IF({{Gewinn23_pflichtig}}=1,"⚠ Verkauf innerhalb 10 Jahren: Gewinn inkl. AfA-Rückrechnung steuerpflichtig.","✔ Verkauf nach Ablauf von 10 Jahren steuerfrei (Notartermin nach dem Stichtag!)."))', None)

    r += 1
    m.section(ws, r, "Prüfungen & Warnhinweise"); r += 1
    C("W_AN", "Anschaffungsnaher Herstellungsaufwand",
      'IF({Verm_Anteil}=0,"– nicht relevant (keine Vermietung)",IF({AN_ueber}=1,"⚠ 15-%-Grenze überschritten: Instandsetzungen der ersten 3 Jahre ("&FIXED({AN_Summe},0)&" €"&" netto > "&FIXED({AN_Grenze},0)&" €"&") werden als Herstellungskosten über die AfA verteilt.","✔ Unter der 15-%-Grenze (Spielraum noch "&FIXED({AN_Grenze}-{AN_Summe},0)&" €"&" netto in 3 Jahren)."))', None)
    C("W_7b", "Sonder-AfA §7b",
      f'IF({{Sonder7b_Wahl}}<>"{JA}","– nicht gewählt",IF({{Sonder7b_ber}}=1,IF({{Horizont}}<10,"⚠ §7b aktiv, aber Vermietung < 10 Jahre → Sonder-AfA wird rückgängig gemacht!","✔ §7b aktiv (Wohnung 10 Jahre entgeltlich vermieten, De-minimis-Grenze beachten)."),"✖ Voraussetzungen nicht erfüllt (Bauantrag-Zeitraum, EH40/QNG, Baukosten ≤ "&FIXED({{P_7b_BKO}},0)&" €/m² oder Erwerb im Fertigstellungsjahr)."))', None)
    C("W_Degr", "Degressive AfA",
      f'IF({{Degr_Wahl}}<>"{JA}","– nicht gewählt",IF({{Degr_aktiv}}=1,"✔ Degressive AfA aktiv (5 % vom Restwert, automatischer Wechsel zu linear).","✖ Voraussetzungen nicht erfüllt (Baubeginn 1.10.2023–30.9.2029 und Erwerb bis Ende des Fertigstellungsjahres)."))', None)
    C("W_RND", "Restnutzungsdauer-Gutachten",
      f'IF({{RND_Gutachten}}<>"{JA}","– nicht gewählt",IF({{RND_aktiv}}=1,"✔ AfA "&FIXED({{AfA_Satz_linear}}*100,2)&" %"&" statt "&FIXED({{AfA_Satz_Std}}*100,2)&" %"&" – Gutachten muss Finanzamt überzeugen (BFH IX R 25/19).","✖ Gutachten-RND ergibt keine höhere AfA."))', None)
    C("W_Inventar", "Inventar",
      'IF({Inventar}>0.15*{Kaufpreis},"⚠ Inventar > 15 % des Kaufpreises – Finanzamt prüft Angemessenheit (Grunderwerbsteuer!).",IF({Inventar}>0,"✔ Inventar separat ausgewiesen: spart "&FIXED({Inventar}*{GrESt_Satz},0)&" €"&" GrESt und wird in "&{Inventar_ND}&" Jahren abgeschrieben.","– kein Inventar ausgewiesen"))', None)
    m.st_post_end = r

    # Jahrestabelle Steuer
    r += 1
    m.section(ws, r, "Steuerliche Jahresübersicht (Auszug aus dem Cashflow-Plan)", ncols=17); r += 1
    cols = [("Jahr", None), ("Einnahmen V+V", "Einnahmen"), ("WK Bewirtschaftung", "WK_Bewirt"),
            ("WK Schuldzinsen", "WK_Zinsen"), ("WK Finanzierungs-NK", "WK_Finanz"), ("AfA Gebäude", "Geb_AfA"),
            ("Sonder-AfA §7b", "S7b"), ("§7i/§7h", "Denkmal_AfA"), ("Inventar-AfA", "Inventar_AfA"),
            ("WK abziehbar", "WK_abz"), ("Einkünfte V+V", "VuV"), ("§10f", "S10f"), ("§35a", "Erm_35a"), ("§35c", "Erm_35c"),
            ("Steuer (+ Last / − Ersparnis)", "Steuerwirkung"), ("Verlustvortrag", "VV")]
    m.subhead(ws, r, [c[0] for c in cols]); r += 1
    r0 = r
    for t in range(1, T_MAX + 1):
        c = ws.cell(row=r, column=2, value=t)
        m.style(c, fmt=F_INT, align="center")
        for i, (_, key) in enumerate(cols[1:]):
            m.view(ws, r, 3 + i, f"INDEX([[{key}]],{t + 1})", F_EUR)
        r += 1
    ws.conditional_formatting.add(f"B{r0}:Q{r - 1}", FormulaRule(formula=[f"$B{r0}>Horizont"], font=Font(color="A6A6A6")))
    m.protect(ws)


def build_vergleich(m):
    ws = m.sheet("vgl", "Kaufen vs. Mieten")
    m.title(ws, "Vermögensvergleich: Immobilie vs. Alternative",
            "Kapitalanlage: gleiches Eigenkapital + jede Zuzahlung fließt stattdessen in einen ETF. "
            "Eigennutzung: Mieter legt EK und die Kostendifferenz an; ist Kaufen günstiger, legt der Käufer die Differenz an.")
    for col, w in zip("BCDEFGHIJKL", (44, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["G"].width = 16
    r = 4

    def C(name, label, tpl, fmt, hint=None, key=False):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, None, key=key)
        if hint:
            m.text(ws, r, 4, hint, size=9, color="404040")
        r += 1

    m.section(ws, r, "Endvermögen nach Steuern am Ende des Horizonts", ncols=11); r += 1
    C("Vgl_Titel", "Vergleich", f'IF({{Modus}}="{KA}","Kapitalanlage-Immobilie vs. ETF-Sparplan","Kaufen (Eigennutzung) vs. Mieten + Differenz anlegen")', None, key=True)
    C("AbgSt_eff", "Abgeltungsteuer inkl. Soli/KiSt", "{P_AbgSt}/(1+{P_AbgSt}*{KiSt_Satz_eff})*(1+{P_Soli_Satz}+{KiSt_Satz_eff})", F_PCT2)
    C("SPB_eff", "Sparer-Pauschbetrag", "{P_SPB}*{Split}", F_EUR)
    C("ETF_netto", "ETF-Rendite nach Steuern (Näherung)", "{ETF_Rendite}*(1-(1-{P_TFS})*{AbgSt_eff})", F_PCT2,
      "Steuer wird erst beim Verkauf am Ende fällig (Vorabpauschale vereinfacht vernachlässigt).")
    C("DepA_N", "Depot Käufer (aus Überschüssen)", "INDEX([[Depot_A]],{Horizont}+1)", F_EUR)
    C("BasA_N", "… Einzahlungen", "INDEX([[Basis_A]],{Horizont}+1)", F_EUR)
    C("Steuer_DepA", "… Steuer bei Verkauf", "MAX(0,({DepA_N}-{BasA_N})*(1-{P_TFS})-{SPB_eff})*{AbgSt_eff}", F_EUR)
    C("Topf_N", "Nachbeleihung: Folgeobjekt-EK / ETF", "INDEX([[NB_Topf]],{Horizont}+1)", F_EUR)
    C("TopfBas_N", "… Einzahlungen", "INDEX([[NB_Topf_Basis]],{Horizont}+1)", F_EUR)
    C("StTopf", "… Steuer (nur ETF)", f'IF({{NB_Verwendung}}="{NB_ETF}",MAX(0,({{Topf_N}}-{{TopfBas_N}})*(1-{{P_TFS}}))*{{AbgSt_eff}},0)', F_EUR)
    C("Konsum_N", "Nachbeleihung: entnommen für Konsum", "INDEX([[NB_Konsum_kum]],{Horizont}+1)", F_EUR)
    C("EndA", "Endvermögen Immobilie (nach Steuern)", "{Verkauf_Netto}+{DepA_N}-{Steuer_DepA}+{Topf_N}-{StTopf}+{Konsum_N}", F_EUR, key=True,
      hint="Netto-Verkaufserlös + angelegte Überschüsse")
    C("DepB_N", "Depot Alternative", "INDEX([[Depot_B]],{Horizont}+1)", F_EUR)
    C("BasB_N", "… Einzahlungen", "INDEX([[Basis_B]],{Horizont}+1)", F_EUR)
    C("Steuer_DepB", "… Steuer bei Verkauf", "MAX(0,({DepB_N}-{BasB_N})*(1-{P_TFS})-{SPB_eff})*{AbgSt_eff}", F_EUR)
    C("EndB", "Endvermögen Alternative (nach Steuern)", "{DepB_N}-{Steuer_DepB}", F_EUR, key=True)
    C("Vorteil", "Vorteil Immobilie (+) / Alternative (−)", "{EndA}-{EndB}", F_EUR, key=True)
    C("Vorteil_Quote", "… in % des eingesetzten EK", "{Vorteil}/MAX(1000,{EK_Einsatz})", F_PCT, key=True)
    C("Vgl_Fazit", "Ergebnis", 'IF({Vorteil}>=0,"Die Immobilie liegt um "&FIXED({Vorteil},0)&" €"&" vorn.","Die Alternative liegt um "&FIXED(-{Vorteil},0)&" €"&" vorn.")', None, key=True)

    r += 1
    m.section(ws, r, "Entwicklung Jahr für Jahr (Depots vor Steuern)", ncols=11); r += 1
    cols = [("Jahr", None), ("Belastung Immobilie p.a. (− = Überschuss)", "_neg_CF"), ("Ausgaben Alternative p.a.", "Alt_Ausgabe"),
            ("Mehrkosten Immobilie", "Diff"), ("Immobilienwert", "Wert"), ("Restschuld", "RS_Summe"),
            ("Depot Käufer", "Depot_A"), ("Depot Alternative", "Depot_B"), ("Vermögen Immobilie", "Verm_A"),
            ("Vermögen Alternative", "Verm_B")]
    m.subhead(ws, r, [c[0] for c in cols]); r += 1
    r0 = r
    for t in range(1, T_MAX + 1):
        c = ws.cell(row=r, column=2, value=t)
        m.style(c, fmt=F_INT, align="center")
        for i, (_, key) in enumerate(cols[1:]):
            f = f"-INDEX([[CF_nach]],{t + 1})" if key == "_neg_CF" else f"INDEX([[{key}]],{t + 1})"
            m.view(ws, r, 3 + i, f, F_EUR)
        r += 1
    ws.conditional_formatting.add(f"B{r0}:K{r - 1}", FormulaRule(formula=[f"$B{r0}>Horizont"], font=Font(color="A6A6A6")))
    m.protect(ws)


def build_dashboard(m):
    ws = m.sheet("kz", "Kennzahlen & Ampel")
    m.title(ws, "Kennzahlen & Ampel", "Dashboard. Schwellwerte für die Ampel auf dem Blatt 'Parameter'.")
    for col, w in zip("BCDEFGHIJ", (46, 22, 14, 14, 10, 10, 4, 40, 16)):
        ws.column_dimensions[col].width = w
    r = 4

    def C(name, label, tpl, fmt, hint=None, key=False, col=3):
        nonlocal r
        m.calc(ws, r, name, label, tpl, fmt, None, key=key, col=col)
        if hint:
            m.text(ws, r, 4, hint, size=9, color="404040")
        r += 1

    # Kopf
    m.section(ws, r, "Fazit", ncols=7); r += 1
    m.fazit_row = r
    r += 3  # Platz für Fazit (wird unten befüllt)

    m.section(ws, r, "Rendite & Cashflow", ncols=7); r += 1
    C("Jahresmiete_Bew", "Jahreskaltmiete (Bewertungsbasis)", f'IF({{Modus}}="{KA}",{{Miete_Start_eff}}*12,{{Miete_Alt_Start}}+{{Miete_Start_eff}}*12)', F_EUR,
      "Eigennutzung: ersparte Miete + Teilvermietung")
    C("Brutto_Rendite", "Bruttomietrendite", "{Jahresmiete_Bew}/{Kaufpreis}", F_PCT2)
    C("Netto_Rendite", "Nettomietrendite", "({Jahresmiete_Bew}*(1-{Leerstand_eff})-{N_Start}-{Instand_pa}-{U_Start}*{Leerstand_eff})/({Kaufpreis}+{NK_Summe})", F_PCT2,
      "nach nicht umlagef. Kosten, Instandhaltung, Ausfall; auf Kaufpreis + NK")
    C("KP_Faktor", "Kaufpreisfaktor", "IF({Jahresmiete_Bew}>0,{Kaufpreis}/{Jahresmiete_Bew},99)", F_NUM1, "Kaufpreis / Jahreskaltmiete")
    C("CF_vor_1", "Cashflow vor Steuern Jahr 1 (p.a.)", "INDEX([[CF_vor]],2)", F_EUR)
    C("CF_vor_1_mtl", "… pro Monat", "{CF_vor_1}/12", F_EUR)
    C("Steuer_1", "Steuerwirkung Jahr 1 (+ Erstattung)", "-INDEX([[Steuerwirkung]],2)", F_EUR)
    C("CF_nach_1", "Cashflow nach Steuern Jahr 1 (p.a.)", "INDEX([[CF_nach]],2)", F_EUR, key=True)
    C("CF_nach_1_mtl", "… pro Monat", "{CF_nach_1}/12", F_EUR, key=True)
    C("Zuzahlung_1", "Monatliche Zuzahlung Jahr 1", "MAX(0,-{CF_nach_1_mtl})", F_EUR, "Eigennutzung: monatliche Wohnkosten netto")
    C("CF_avg_mtl", "Ø Cashflow n. St. pro Monat über Horizont", "SUMPRODUCT(([[t]]>=1)*([[t]]<={Horizont})*[[CF_nach]])/{Horizont}/12", F_EUR)
    C("Steuer_Summe", "Steuerersparnis (+) / -last (−) über Horizont", "-SUMPRODUCT(([[t]]>=1)*([[t]]<={Horizont})*[[Steuerwirkung]])", F_EUR, "ohne Steuer auf den Verkauf")
    C("EK_Rendite_1", "Eigenkapitalrendite Jahr 1", "({CF_nach_1}+INDEX([[Tilgung_Summe]],2))/MAX(1000,{EK_Einsatz})", F_PCT2, "(CF n. St. + Tilgung) / EK")
    for gi, g in enumerate(("0.05", "0", "-0.1", "0.2", "-0.3", "-0.6")):
        irr = f"IRR([[IRR_Flow]],{g})"
        tpl = f'IFERROR(IF(AND({irr}>-0.99,{irr}<2),{irr},""),"")'
        m.nodes[f"IRR_Kand{gi}"] = {"tpl": tpl, "kind": "calc"}
        c = ws.cell(row=r, column=10 + gi, value=m.to_main(tpl))
        c.font = fnt(size=7, color="BFBFBF")
        m.define(f"IRR_Kand{gi}", ws, f"${get_column_letter(10 + gi)}${r}")
    pick = '"n. b."'
    for gi in reversed(range(6)):
        pick = f"IF(ISNUMBER({{IRR_Kand{gi}}}),{{IRR_Kand{gi}}},{pick})"
    C("IRR_Wert", "IRR nach Steuern (inkl. Verkauf)", pick, F_PCT2, "Eigennutzung: inkl. ersparter Miete", key=True)
    C("DSCR_1", "Kapitaldienstdeckung Jahr 1 (DSCR)",
      "IF(INDEX([[Kapitaldienst]],2)>0,(INDEX([[Miete_Ist]],2)-INDEX([[U_Traeger]],2)-INDEX([[N_Kosten]],2)-INDEX([[Instand]],2))/INDEX([[Kapitaldienst]],2),99)", F_NUM2,
      "Mietüberschuss / (Zins + Tilgung)")
    C("BE_V_Jahr", "Break-even Vermögen (Jahr)", 'IF(MIN([[BE_V]])>=999,"nicht erreicht",MIN([[BE_V]]))', F_INT,
      "KA: Liquidationswert inkl. kum. CF ≥ EK; EN: Kaufen ≥ Mieten")
    C("BE_CF_Jahr", "Break-even Cashflow (erstes Jahr CF n. St. ≥ 0)", 'IF(MIN([[BE_CF]])>=999,"nicht erreicht",MIN([[BE_CF]]))', F_INT)
    C("Belastungsquote", "Belastungsquote", f'IF({{Netto_Einkommen}}>0,IF({{Modus}}="{KA}",{{Zuzahlung_1}},-{{CF_nach_1_mtl}})/{{Netto_Einkommen}},0)', F_PCT,
      "KA: Zuzahlung / Netto; EN: Wohnkosten / Netto")

    r += 1
    m.section(ws, r, "Vermögen", ncols=7); r += 1
    C("Vermoegen_N", "Vermögen am Ende (Wert − Schuld + kum. CF)", "INDEX([[Vermoegen]],{Horizont}+1)", F_EUR, "Definition laut Vorgabe, vor Verkaufskosten/Steuer", key=True)
    C("CF_kum_N", "davon kumulierter Cashflow", "INDEX([[CF_kum]],{Horizont}+1)", F_EUR)
    C("Vermoegen_Zuwachs", "Vermögenszuwachs ggü. eingesetztem EK", "{Vermoegen_N}-{EK_Einsatz}", F_EUR)
    C("EndA_v", "Endvermögen Immobilie nach Steuern", "{EndA}", F_EUR)
    C("EndB_v", "Endvermögen Alternative nach Steuern", "{EndB}", F_EUR)
    C("Vorteil_v", "Vorteil Immobilie ggü. Alternative", "{Vorteil}", F_EUR, key=True)

    # Ampel
    r += 1
    m.section(ws, r, "Ampel – Kriterien", ncols=7); r += 1
    m.subhead(ws, r, ["Kriterium", "Wert", "grün", "gelb", "Punkte", "Gewicht"]); r += 1
    crit = [
        ("IRR", "IRR nach Steuern", "{IRR_Wert}", "P_Amp_IRR_gr", "P_Amp_IRR_ge", True, f'{{Modus}}="{KA}"', "P_Amp_W_IRR", F_PCT2),
        ("Vort", "Vermögensvorteil ggü. Alternative (% EK)", "{Vorteil_Quote}", "P_Amp_Vort_gr", "P_Amp_Vort_ge", True, "TRUE", "P_Amp_W_Vort", F_PCT),
        ("CF", "Cashflow n. St. Jahr 1 (€/Monat)", "{CF_nach_1_mtl}", "P_Amp_CF_gr", "P_Amp_CF_ge", True, f'{{Modus}}="{KA}"', "P_Amp_W_CF", F_EUR),
        ("KPF", "Kaufpreisfaktor", "{KP_Faktor}", "P_Amp_KPF_gr", "P_Amp_KPF_ge", False, "TRUE", "P_Amp_W_KPF", F_NUM1),
        ("DSCR", "Kapitaldienstdeckung", "{DSCR_1}", "P_Amp_DSCR_gr", "P_Amp_DSCR_ge", True, f'{{Modus}}="{KA}"', "P_Amp_W_DSCR", F_NUM2),
        ("BQKA", "Belastungsquote (Kapitalanlage)", "{Belastungsquote}", "P_Amp_BQKA_gr", "P_Amp_BQKA_ge", False, f'{{Modus}}="{KA}"', "P_Amp_W_BQ", F_PCT),
        ("BQEN", "Wohnkostenquote (Eigennutzung)", "{Belastungsquote}", "P_Amp_BQEN_gr", "P_Amp_BQEN_ge", False, f'{{Modus}}="{EN}"', "P_Amp_W_BQ", F_PCT),
    ]
    amp_rows = []
    for key, label, val, gr, ge, higher, appl, w, fmt in crit:
        m.label(ws, r, label)
        m.view(ws, r, 3, val, fmt)
        m.view(ws, r, 4, "{" + gr + "}", fmt)
        m.view(ws, r, 5, "{" + ge + "}", fmt)
        op = ">=" if higher else "<="
        m.calc(ws, r, f"Amp_P_{key}", None,
               f'IF(NOT({appl}),0,IF(NOT(ISNUMBER({val})),0,IF({val}{op}{{{gr}}},2,IF({val}{op}{{{ge}}},1,0))))', F_INT, col=6, write_label=False)
        m.calc(ws, r, f"Amp_W_{key}", None, f"IF({appl},{{{w}}},0)", F_NUM1, col=7, write_label=False)
        amp_rows.append(r)
        r += 1
    ws.column_dimensions["G"].width = 9
    rng = f"F{amp_rows[0]}:F{amp_rows[-1]}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=["2"], fill=fill(C_GREEN)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=["1"], fill=fill(C_YELLOW)))
    rng2 = f"F{amp_rows[0]}:F{amp_rows[-1]}"
    ws.conditional_formatting.add(rng2, FormulaRule(formula=[f"AND(F{amp_rows[0]}=0,G{amp_rows[0]}>0)"], fill=fill(C_RED)))
    keys = [c[0] for c in crit]
    score_tpl = "(" + "+".join(f"{{Amp_P_{k}}}*{{Amp_W_{k}}}" for k in keys) + ")/MAX(0.001," + "+".join(f"{{Amp_W_{k}}}" for k in keys) + ")"
    C("Score", "Gesamtpunktzahl (0 = rot … 2 = grün)", score_tpl, F_NUM2, key=True)
    C("Fazit", "Fazit",
      'IF(AND({Score}>={P_Amp_Score_gr},{Amp_P_Vort}>0),"LOHNT SICH",IF(AND({Score}>={P_Amp_Score_ge},{Amp_P_Vort}+{Score}>0),"GRENZWERTIG","LOHNT SICH NICHT"))', None, key=True,
      hint="K.o.: Liegt die Alternative deutlich vorn (rot), ist höchstens 'grenzwertig' möglich.")

    # Fazit-Kopf
    fr = m.fazit_row
    ws.merge_cells(start_row=fr, start_column=2, end_row=fr + 1, end_column=3)
    c = ws.cell(row=fr, column=2, value="=Fazit")
    c.font = fnt(bold=True, size=20)
    c.alignment = Alignment(horizontal="center", vertical="center")
    for rr in (fr, fr + 1):
        for cc in (2, 3):
            ws.cell(row=rr, column=cc).border = BORDER
    ws.conditional_formatting.add(f"B{fr}:C{fr + 1}", FormulaRule(formula=['Fazit="LOHNT SICH"'], fill=fill(C_GREEN), font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(f"B{fr}:C{fr + 1}", FormulaRule(formula=['Fazit="GRENZWERTIG"'], fill=fill(C_YELLOW), font=Font(color="9C5700", bold=True)))
    ws.conditional_formatting.add(f"B{fr}:C{fr + 1}", FormulaRule(formula=['Fazit="LOHNT SICH NICHT"'], fill=fill(C_RED), font=Font(color="9C0006", bold=True)))
    m.view(ws, fr, 4, '{Modus}&" · Szenario "&{Szenario}&" · "&{Horizont}&" Jahre"', None, bold=True)
    ws.merge_cells(start_row=fr, start_column=4, end_row=fr, end_column=7)
    m.view(ws, fr + 1, 4, '"IRR: "&IF(ISNUMBER({IRR_Wert}),FIXED({IRR_Wert}*100,1)&" %","n. b.")&" · CF n. St. Jahr 1: "&FIXED({CF_nach_1_mtl},0)&" €"&"/Monat · "&{Vgl_Fazit}', None, wrap=True)
    ws.merge_cells(start_row=fr + 1, start_column=4, end_row=fr + 1, end_column=7)
    ws.row_dimensions[fr + 1].height = 30

    # Warnungen
    r += 1
    m.section(ws, r, "Hinweise & Warnungen", ncols=7); r += 1
    for w in ("Fin_Hinweis", "W_AN", "W_Verbilligt", "W_Spek", "W_DreiObj", "W_7b", "W_Degr", "W_RND", "W_Inventar"):
        m.view(ws, r, 2, "{" + w + "}", None, wrap=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 28
        r += 1

    # Diagramme
    cf = m.sheets["cf"]
    r0 = m.eng_row0 + 1
    r1 = m.eng_row0 + T_MAX

    from openpyxl.utils import column_index_from_string as cis
    cats = Reference(cf, min_col=cis(m.eng_col["t"]), min_row=r0, max_row=r1)

    ch = LineChart()
    ch.title = "Vermögen: Immobilie vs. Alternative (vor Steuern)"
    ch.y_axis.title = "€"
    ch.x_axis.title = "Jahr"
    ch.height, ch.width = 8, 17
    for key, name in (("Verm_A", "Immobilie"), ("Verm_B", "Alternative"), ("Vermoegen", "Wert − Schuld + kum. CF")):
        s = Reference(cf, min_col=cis(m.eng_col[key]), min_row=r0, max_row=r1)
        ch.add_data(s, titles_from_data=False)
        ch.series[-1].tx = None
    from openpyxl.chart.series import SeriesLabel
    for sname, s in zip(("Immobilie", "Alternative", "Wert − Schuld + kum. CF"), ch.series):
        s.tx = SeriesLabel(v=sname)
    ch.set_categories(cats)
    ws.add_chart(ch, "I4")

    ch2 = BarChart()
    ch2.title = "Cashflow vor / nach Steuern p.a."
    ch2.height, ch2.width = 8, 17
    for key, name in (("CF_vor", "vor Steuern"), ("CF_nach", "nach Steuern")):
        ch2.add_data(Reference(cf, min_col=cis(m.eng_col[key]), min_row=r0, max_row=r1), titles_from_data=False)
    for sname, s in zip(("vor Steuern", "nach Steuern"), ch2.series):
        s.tx = SeriesLabel(v=sname)
    ch2.set_categories(cats)
    ws.add_chart(ch2, "I21")

    ch3 = LineChart()
    ch3.title = "Immobilienwert vs. Restschuld"
    ch3.height, ch3.width = 8, 17
    for key in ("Wert", "RS_Summe"):
        ch3.add_data(Reference(cf, min_col=cis(m.eng_col[key]), min_row=r0, max_row=r1), titles_from_data=False)
    for sname, s in zip(("Immobilienwert", "Restschuld"), ch3.series):
        s.tx = SeriesLabel(v=sname)
    ch3.set_categories(cats)
    ws.add_chart(ch3, "I38")
    m.text(ws, 55, 9, "Diagramme zeigen 40 Jahre; ausgewertet wird bis zum gewählten Horizont.", italic=True, size=8, color="7F7F7F")
    m.protect(ws)


# ====================================================================== Schattenrechnung
def tpl_refs(tpl):
    scal = set(re.findall(r"\{(\w+)\}", tpl))
    eng = set(re.findall(r"\[\[(\w+)\]\]", tpl)) | set(re.findall(r"\[(\w+)(?:@-1)?\]", tpl))
    return scal, eng


class Shadow:
    def __init__(self, m):
        self.m = m
        self.ecols = {c["key"]: c for c in m.engine}
        self.sdeps, self.edeps = {}, {}
        for n, d in m.nodes.items():
            if d["tpl"]:
                self.sdeps[n] = tpl_refs(d["tpl"])
        for k, c in self.ecols.items():
            s, e = set(), set()
            for t in (c["tpl"], c["init"]):
                if t:
                    a, b = tpl_refs(t)
                    s |= a
                    e |= b
            self.edeps[k] = (s, e)

    def plan(self, overrides, outputs):
        need_s, need_e = set(), set()
        stack = [("s", o) for o in outputs]
        while stack:
            kind, n = stack.pop()
            if kind == "s":
                if n in need_s:
                    continue
                need_s.add(n)
                if n in overrides:
                    continue
                if n in self.sdeps:
                    s, e = self.sdeps[n]
                    stack += [("s", x) for x in s] + [("e", x) for x in e]
            else:
                if n in need_e:
                    continue
                need_e.add(n)
                s, e = self.edeps[n]
                stack += [("s", x) for x in s] + [("e", x) for x in e]
        memo = {}

        def dyn(n):
            if n in memo:
                return memo[n]
            memo[n] = False
            if n in overrides:
                memo[n] = True
            elif n in self.sdeps:
                s, e = self.sdeps[n]
                memo[n] = bool(e) or any(dyn(x) for x in s)
            return memo[n]

        local = [n for n in need_s if dyn(n)]
        return sorted(local), [c["key"] for c in self.m.engine if c["key"] in need_e]

    def emit(self, ws, top, title, overrides, outputs, out_prefix):
        local, cols = self.plan(overrides, outputs)
        smap = {n: f"$B${top + 2 + i}" for i, n in enumerate(local)}
        cmap = {k: get_column_letter(4 + i) for i, k in enumerate(cols)}
        r0 = top + 2  # Jahr 0

        def conv(tpl, row=None):
            s = re.sub(r"\[\[(\w+)\]\]", lambda mm: f"${cmap[mm.group(1)]}${r0}:${cmap[mm.group(1)]}${r0 + T_MAX}", tpl)
            if row is not None:
                s = re.sub(r"\[(\w+)@-1\]", lambda mm: f"{cmap[mm.group(1)]}{row - 1}", s)
                s = re.sub(r"\[(\w+)\]", lambda mm: f"{cmap[mm.group(1)]}{row}", s)
            s = re.sub(r"\{(\w+)\}", lambda mm: smap.get(mm.group(1), mm.group(1)), s)
            assert "[" not in s and "{" not in s, s
            return "=" + s

        ws.cell(row=top, column=1, value=title).font = fnt(bold=True)
        for i, n in enumerate(local):
            rr = top + 2 + i
            ws.cell(row=rr, column=1, value=n).font = fnt(size=8, color="7F7F7F")
            f = overrides[n] if n in overrides else conv(self.m.nodes[n]["tpl"])
            ws.cell(row=rr, column=2, value=f)
        for k in cols:
            c = self.ecols[k]
            L = cmap[k]
            ws[f"{L}{top + 1}"] = k
            for t in range(0, T_MAX + 1):
                rr = r0 + t
                if c["const"]:
                    v = t
                elif t == 0:
                    v = conv(c["init"], rr) if c["init"] is not None else None
                else:
                    v = conv(c["tpl"], rr)
                ws[f"{L}{rr}"] = v
        for o in outputs:
            self.m.define(f"{out_prefix}_{o}", ws, smap[o])
        return top + max(len(local), T_MAX + 1) + 4


def build_sensitivity(m):
    sh = Shadow(m)
    ws_s = m.sheet("sens", "_Sens")
    ws_s.sheet_state = "hidden"
    ws = m.sheets["sz"]
    r = m.sz_row + 1
    top = 1
    outs = ["IRR_Wert", "Vorteil", "CF_nach_1_mtl", "Vermoegen_N", "EndA", "EndB", "Fazit", "CF_nach_1", "Steuer_Summe", "Score"]

    # Szenariovergleich
    for i, sz in enumerate(["Base", "Best", "Worst"]):
        top = sh.emit(ws_s, top, f"Szenario {sz}", {"Szen_Idx": f"={i + 1}"}, outs, f"Szen_{sz}")
    m.section(ws, r, "Szenariovergleich (alle drei gleichzeitig gerechnet)", ncols=8); r += 1
    m.subhead(ws, r, ["Kennzahl", "Base", "Best", "Worst"]); r += 1
    for o, label, fmt in (("IRR_Wert", "IRR nach Steuern", F_PCT2), ("CF_nach_1_mtl", "Cashflow n. St. Jahr 1 / Monat", F_EUR),
                          ("Steuer_Summe", "Steuerersparnis über Horizont", F_EUR), ("Vermoegen_N", "Vermögen am Ende (Wert − Schuld + kum. CF)", F_EUR),
                          ("EndA", "Endvermögen Immobilie n. St.", F_EUR), ("EndB", "Endvermögen Alternative n. St.", F_EUR),
                          ("Vorteil", "Vorteil Immobilie", F_EUR), ("Score", "Ampel-Punkte", F_NUM2), ("Fazit", "Fazit", None)):
        m.label(ws, r, label)
        for j, sz in enumerate(["Base", "Best", "Worst"]):
            c = ws.cell(row=r, column=3 + j, value=f"=Szen_{sz}_{o}")
            m.style(c, fmt=fmt, bold=o == "Fazit")
        r += 1
    m.text(ws, r, 2, "Hinweis: Base nutzt deine Eingaben unverändert; Best/Worst wenden die Abweichungen oben an.", italic=True, size=9, color="7F7F7F")
    r += 2

    # Kennzahl-Auswahl
    m.section(ws, r, "Sensitivitätsraster", ncols=8); r += 1
    m.label(ws, r, "Angezeigte Kennzahl in den Rastern", None, "Wähle, welche Kennzahl in beiden Rastern angezeigt wird.", hint_col=8)
    m.inp(ws, r, "Sens_Kennzahl", None, KENNZAHLEN[0], None, ("list", "Liste_Kennzahl"),
          "Wähle, welche Kennzahl in beiden Rastern angezeigt wird.", write_label=False)
    r += 1
    m.calc(ws, r, "Sens_KZ_Idx", "Index", "MATCH({Sens_Kennzahl},{Liste_Kennzahl},0)", F_INT)
    r += 2
    out4 = ["IRR_Wert", "Vorteil", "CF_nach_1_mtl", "Vermoegen_N"]

    def grid(tag, title, row_name, row_label, row_vals, row_fmt, col_name, col_label, col_vals, col_fmt, ovf):
        nonlocal r, top
        m.section(ws, r, title, ncols=8); r += 1
        m.text(ws, r, 2, f"Zeilen: {row_label} · Spalten: {col_label} (gelbe Achsenwerte änderbar)", italic=True, size=9)
        r += 1
        hdr = r
        c = ws.cell(row=hdr, column=2, value=f"{row_label} ↓ / {col_label} →")
        m.style(c, bold=True, fillc=C_SUB, wrap=True)
        for j, v in enumerate(col_vals):
            m.inp(ws, hdr, f"Sens{tag}_{col_name}_{j + 1}", None, v, col_fmt,
                  ("dec", -1, 1, "Bitte einen Wert zwischen −100 % und +100 % eingeben."), None, col=3 + j, write_label=False)
        for i, v in enumerate(row_vals):
            rr = hdr + 1 + i
            m.inp(ws, rr, f"Sens{tag}_{row_name}_{i + 1}", None, v, row_fmt,
                  ("dec", -1, 1, "Bitte einen Wert zwischen −100 % und +100 % eingeben."), None, col=2, write_label=False)
            for j in range(len(col_vals)):
                pref = f"Sens_{tag}_{i + 1}_{j + 1}"
                top = sh.emit(ws_s, top, pref, ovf(i + 1, j + 1), out4, pref)
                f = "=CHOOSE(Sens_KZ_Idx," + ",".join(f"{pref}_{o}" for o in out4) + ")"
                cell = ws.cell(row=rr, column=3 + j, value=f)
                m.style(cell, fmt="General")
        rng = f"C{hdr + 1}:{get_column_letter(2 + len(col_vals))}{hdr + len(row_vals)}"
        # Zahlformat je Kennzahl: Standard Euro, bei IRR per bedingter Formatierung Prozent
        for rr in range(hdr + 1, hdr + 1 + len(row_vals)):
            for cc in range(3, 3 + len(col_vals)):
                ws.cell(row=rr, column=cc).number_format = '#,##0 "€";[Red]-#,##0 "€"'
        ws.conditional_formatting.add(rng, Rule(type="expression", formula=["Sens_KZ_Idx=1"],
                                                dxf=DifferentialStyle(numFmt=NumberFormat(numFmtId=200, formatCode="0.0%"))))
        mid = f"{get_column_letter(3 + len(col_vals) // 2)}{hdr + 1 + len(row_vals) // 2}"
        ws[mid].border = Border(left=Side(style="medium"), right=Side(style="medium"), top=Side(style="medium"), bottom=Side(style="medium"))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER(C{hdr + 1}),C{hdr + 1}<0)"], font=Font(color="C00000")))
        r = hdr + len(row_vals) + 2

    grid("A", "Raster 1: Zinsniveau × Kaufpreis", "Zins", "Zinsänderung (alle Darlehen, Δ Prozentpunkte)", [-0.01, -0.005, 0, 0.005, 0.01], F_PCT2,
         "KP", "Kaufpreisänderung", [-0.10, -0.05, 0, 0.05, 0.10], F_PCT,
         lambda i, j: {"Sens_Zins_Delta": f"=SensA_Zins_{i}", "Kaufpreis": f"=Kaufpreis*(1+SensA_KP_{j})"})
    grid("B", "Raster 2: Miete × Wertsteigerung", "Miete", "Mietänderung (Start)", [-0.10, -0.05, 0, 0.05, 0.10], F_PCT,
         "Wert", "Wertsteigerung (Δ p.a.)", [-0.015, -0.0075, 0, 0.0075, 0.015], F_PCT2,
         lambda i, j: {"Miete_Start_eff": f"=Miete_Start_eff*(1+SensB_Miete_{i})",
                       "Miete_Alt_Start": f"=Miete_Alt_Start*(1+SensB_Miete_{i})",
                       "Wertsteig_eff": f"=Wertsteig_eff+SensB_Wert_{j}"})
    m.text(ws, r, 2, "Formatierung: Werte < 1 werden als Prozent (IRR), sonst als Euro angezeigt. Mitte (fett umrandet) = aktuelle Eingaben.",
           italic=True, size=9, color="7F7F7F")
    m.protect(ws)
    m.protect(ws_s)





# ====================================================================== Checkliste & Start
def build_checkliste(m):
    ws = m.sheet("chk", "Steuer-Checkliste")
    m.title(ws, "Steuer-Checkliste – alle legalen Hebel",
            "Status wird automatisch aus deinen Eingaben ermittelt. ✔ greift · ✖ greift nicht · ⚠ prüfen · ℹ Hinweis (nicht gerechnet).")
    for col, w in zip("BCDEF", (40, 22, 70, 30, 16)):
        ws.column_dimensions[col].width = w
    r = 4
    m.subhead(ws, r, ["Hebel", "Status im aktuellen Fall", "Erklärung", "Rechtsgrundlage", "Wirkung (Info)"]); r += 1
    V = f'{{Verm_Anteil}}>0'
    items = [
        ("A  Kauf & Kaufpreisaufteilung", None, None, None, None),
        ("Kaufpreisaufteilung im Kaufvertrag", f'IF({{Aufteilung_Modus}}="{AUF_MAN}","✔ manuelle Aufteilung genutzt","⚠ prüfen: Gebäudeanteil vertraglich festlegen")',
         "Ein hoher, aber plausibler Gebäudeanteil im Kaufvertrag erhöht die AfA. Finanzamt folgt dem Vertrag, solange er die Werte nicht nur zum Schein verteilt.", "BFH IX R 26/19; §6f EStG-E (JStG 2026)", "{AK_Gebaeude}"),
        ("Inventar separat ausweisen", 'IF({Inventar}>0,"✔ greift","✖ kein Inventar ausgewiesen")',
         "Einbauküche, Markise, Sauna etc. separat bewerten: keine GrESt, Abschreibung über ca. 10 Jahre.", "§2 GrEStG; §7 Abs. 1 EStG", "{Inventar}*{GrESt_Satz}"),
        ("Instandhaltungsrücklage", '"ℹ Hinweis"', "Der anteilige Rücklagenbetrag mindert NICHT die Grunderwerbsteuer (BFH II R 49/17).", "BFH 16.9.2020, II R 49/17", None),
        ("Finanzierungsnebenkosten sofort abziehen", f'IF({V},"✔ greift","✖ nur bei Vermietung")',
         "Notar/Grundbuch für die Grundschuld, Bereitstellungszinsen, Schätzkosten der Bank: sofort Werbungskosten statt AK.", "§9 Abs. 1 Nr. 1 EStG", "{NK_Finanzierung}"),
        ("Disagio bis 5 % sofort abziehen", f'IF(AND({V},{{Disagio_Summe}}>0),"✔ greift","– nicht genutzt")',
         "Marktübliches Disagio (≤ 5 %, Zinsbindung ≥ 5 J.) ist im Zahlungsjahr abziehbar – zieht Steuerersparnis vor (Zinsnachteil beachten).", "§11 Abs. 2 S. 4 EStG; BMF 20.10.2003", "{Disagio_Summe}"),
        ("B  Abschreibung", None, None, None, None),
        ("Lineare AfA nach Baujahr", f'IF({V},"✔ "&{{AfA_Art}},"– nur bei Vermietung")', "2,5 % (vor 1925), 2 % (1925–2022), 3 % (ab 2023) – wird automatisch gewählt.", "§7 Abs. 4 EStG", "INDEX([[Geb_AfA]],2)"),
        ("Kürzere Restnutzungsdauer (Gutachten)", f'IF({{RND_aktiv}}=1,"✔ greift",IF({V},"⚠ prüfen: lohnt bei Altbauten oft","– nur bei Vermietung"))',
         "Gutachten eines Sachverständigen kann z. B. 40 statt 50 Jahre belegen → 2,5 % statt 2 %. Gutachterkosten sind Werbungskosten.", "§7 Abs. 4 S. 2 EStG; BFH IX R 25/19", None),
        ("Degressive AfA 5 % (Neubau)", 'IF({Degr_aktiv}=1,"✔ greift",IF({Degr_ber}=1,"⚠ zulässig, aber nicht gewählt","✖ nicht zulässig"))',
         "Wohngebäude mit Baubeginn 1.10.2023–30.9.2029; Erwerb bis Ende des Fertigstellungsjahres. Wechsel zu linear jederzeit.", "§7 Abs. 5a EStG", None),
        ("Sonder-AfA §7b", 'IF({Sonder7b_aktiv}=1,"✔ greift",IF({Sonder7b_ber}=1,"⚠ zulässig, aber nicht gewählt","✖ nicht zulässig"))',
         "4 × 5 % zusätzlich für neue Mietwohnungen (EH40 + QNG, Baukosten ≤ 5.200 €/m², BMG ≤ 4.000 €/m², 10 J. Vermietung).", "§7b EStG", None),
        ("Denkmal-AfA §7i / Sanierungsgebiet §7h", f'IF(AND({V},{{Denkmal_BMG}}>0),"✔ greift",IF({{Denkmal_BMG}}>0,"– Eigennutzung: siehe §10f","✖ kein Denkmal"))',
         "Begünstigte Sanierungskosten: 8 × 9 % + 4 × 7 % = 100 % in 12 Jahren. Bescheinigung VOR Beginn der Maßnahme abstimmen.", "§7i, §7h, §11a/b EStG", None),
        ("§10f Eigennutzung Denkmal", f'IF(AND({{Modus}}="{EN}",{{Denkmal_BMG}}>0),"✔ greift","– nicht relevant")', "Selbst genutztes Denkmal: 10 × 9 % als Sonderausgabe.", "§10f EStG", None),
        ("Inventar-AfA", f'IF(AND({V},{{Inventar}}>0),"✔ greift","– nicht relevant")', "Einbauküche etc. über Nutzungsdauer abschreiben (GWG bis 800 € netto sofort).", "§7 Abs. 1, §6 Abs. 2 EStG", None),
        ("C  Erhaltung & Modernisierung", None, None, None, None),
        ("15-%-Grenze anschaffungsnaher Aufwand", 'IF({AN_ueber}=1,"⚠ überschritten – Aufwand wird aktiviert",IF({Verm_Anteil}>0,"✔ eingehalten","– nicht relevant"))',
         "Instandsetzungen in den ersten 3 Jahren > 15 % der Gebäude-AK (netto) werden zu Herstellungskosten. Größere Maßnahmen ggf. ins 4. Jahr schieben.", "§6 Abs. 1 Nr. 1a EStG", "{AN_Grenze}-{AN_Summe}"),
        ("Erhaltung statt Herstellung (Standardhebung vermeiden)", '"ℹ Hinweis"',
         "Herstellungsaufwand liegt vor bei Erweiterung oder Hebung von mind. 3 der 4 Kernbereiche (Heizung, Sanitär, Elektro, Fenster). Maßnahmen zeitlich staffeln.", "BMF 18.7.2003; R 21.1 EStR", None),
        ("Erhaltungsaufwand verteilen (§82b EStDV)", f'IF(AND({V},{{Verteilung_eff}}>1),"✔ greift ("&{{Verteilung_eff}}&" Jahre)",IF({V},"⚠ prüfen: bei hohem Einmalaufwand sinnvoll","– nur bei Vermietung"))',
         "Größere Erhaltungsaufwände auf 2–5 Jahre verteilen glättet die Progression (sinnvoll, wenn das zvE in Folgejahren höher ist).", "§82b EStDV", None),
        ("Leerstand: Kosten weiter abziehen", f'IF({V},"✔ greift (bei Vermietungsabsicht)","– nur bei Vermietung")', "Bei ernsthafter Vermietungsabsicht (Inserate dokumentieren) bleiben alle Kosten abziehbar.", "BFH IX R 3/13", None),
        ("D  Finanzierung", None, None, None, None),
        ("Schuldzinsen: Fremdkapital in die Vermietung", f'IF({{Modus}}="{KA}","✔ Zinsen voll abziehbar","ℹ Eigennutzung: Zinsen nicht abziehbar")',
         "Wer selbst nutzt UND vermietet: Eigenkapital in die selbst genutzte, Kredite in die vermietete Immobilie (Zinsen nur dort abziehbar).", "§9 Abs. 1 Nr. 1 EStG", "INDEX([[WK_Zinsen]],2)"),
        ("Gezielte Darlehenszuordnung bei Teilvermietung", f'IF(AND({{Modus}}="{EN}",{{Verm_Anteil}}>0),IF({{Teilverm_Zuordnung}}="{JA}","✔ greift","⚠ prüfen: separates Darlehen + getrennte Kaufpreiszahlung"),"– nicht relevant")',
         "Separates Darlehen, Kaufpreis für den vermieteten Teil gesondert ausweisen und davon zahlen → Zinsen voll beim vermieteten Teil.", "BFH IX R 18/97", None),
        ("Nachbeleihung (später erneut belasten)", f'IF({{NB_an}}=1,IF({{NB_Abzug_Quote}}>0,"✔ Zinsen abziehbar ("&FIXED({{NB_Abzug_Quote}}*100,0)&" %"&")","⚠ Zinsen NICHT abziehbar (Verwendung privat/ETF)"),"– nicht eingeplant")',
         "Abzug folgt der Verwendung: Modernisierung des Mietobjekts oder EK für weiteres Mietobjekt → abziehbar. ETF (Abgeltungsteuer) oder Konsum → nicht abziehbar. Darlehen sauber getrennt führen.", "§9 Abs. 1 Nr. 1, §20 Abs. 9 EStG", None),
        ("Endfällig + Tilgungsersatz statt Tilgung", '"ℹ Hinweis"',
         "Tilgungsaussetzung (z. B. ETF-Sparplan) erhöht dauerhaft die abziehbaren Zinsen; Risiko: Rendite des Tilgungsersatzes. Im Modell: Tilgung 0 % + ETF-Vergleich.", "§9 Abs. 1 Nr. 1 EStG", None),
        ("Nachträgliche Schuldzinsen nach Verkauf", '"ℹ Hinweis"', "Reicht der Verkaufserlös nicht zur Tilgung, bleiben Restzinsen Werbungskosten.", "BFH IX R 67/10, IX R 45/13", None),
        ("KfW-Tilgungszuschuss", 'IF({Kredit1_TZ}+{Kredit2_TZ}+{Kredit3_TZ}+{Kredit4_TZ}>0,"⚠ prüfen: steuerliche Behandlung","– nicht genutzt")',
         "Zuschüsse mindern bei Vermietung die AK/HK bzw. den Erhaltungsaufwand (im Modell vereinfacht als Einnahme). Bei Eigennutzung steuerneutral.", "R 21.5 EStR", None),
        ("E  Laufende Werbungskosten", None, None, None, None),
        ("Alle Werbungskosten erfassen", f'IF({V},"✔ greift","– nur bei Vermietung")',
         "Fahrten (0,30 €/km), Kontoführung, Steuerberatung, Porto/Telefon, Inserate, Rechtsschutz, Haus & Grund, Mietverwaltung, Fachliteratur.", "§9 EStG", "{S_Start}"),
        ("Freibetrag im Lohnsteuerabzug", f'IF(AND({V},INDEX([[VuV]],2)<0),"⚠ prüfen: Liquidität sofort statt mit Steuerbescheid","– nicht relevant")',
         "Verluste aus V+V als Freibetrag (ELStAM) eintragen lassen → monatlich weniger Lohnsteuer statt Erstattung im Folgejahr.", "§39a Abs. 1 Nr. 5 EStG", "-INDEX([[Steuerwirkung]],2)/12"),
        ("Verlustverrechnung mit anderen Einkünften", f'IF(AND({V},INDEX([[VuV]],2)<0),"✔ greift (Verluste mindern Steuer)","– aktuell kein Verlust")',
         "Verluste aus V+V werden mit Lohn etc. verrechnet; nicht genutzte Verluste werden vorgetragen (im Modell berücksichtigt).", "§2 Abs. 3, §10d EStG", None),
        ("F  Vermietung an Angehörige", None, None, None, None),
        ("66-%-Grenze verbilligte Vermietung", f'IF({{Angehoerige}}="{JA}",{{W_Verbilligt}},"– nicht relevant")',
         "Mindestens 66 % der ortsüblichen Warmmiete → voller Abzug. 50–66 %: Totalüberschussprognose. < 50 %: anteilig.", "§21 Abs. 2 EStG", None),
        ("Mietvertrag mit Angehörigen fremdüblich", f'IF({{Angehoerige}}="{JA}","⚠ prüfen","– nicht relevant")',
         "Schriftlicher Vertrag, tatsächliche Zahlung, Nebenkostenabrechnung wie unter Fremden – sonst kein Abzug. Vermietung an unterhaltsberechtigte Kinder möglich.", "BFH IX R 30/98, IX R 46/12", None),
        ("G  Verkauf", None, None, None, None),
        ("Spekulationsfrist 10 Jahre", f'IF({{Gewinn23_pflichtig}}=1,IF({{Verm_Anteil}}>0,"⚠ Verkauf innerhalb der Frist","✔ Eigennutzung steuerfrei"),"✔ steuerfrei")',
         "Kapitalanlage: Verkauf erst nach 10 Jahren (Datum notarieller Vertrag!). Veräußerungsgewinn inkl. AfA-Rückrechnung, Freigrenze 1.000 €.", "§23 Abs. 1 Nr. 1 EStG", "{Steuer_Verkauf}"),
        ("Eigennutzung vor Verkauf", f'IF({{Modus}}="{EN}","✔ greift","ℹ Hinweis")',
         "Steuerfrei auch innerhalb 10 Jahren, wenn im Verkaufsjahr und den 2 Vorjahren selbst genutzt (Kalenderjahre, mittleres Jahr vollständig). Strategie: vor Verkauf selbst einziehen.", "§23 Abs. 1 Nr. 1 S. 3 EStG", None),
        ("Drei-Objekt-Grenze", "{W_DreiObj}", "Mehr als 3 Objekte binnen 5 Jahren nach Kauf/Bau verkauft → gewerblicher Grundstückshandel (GewSt, keine Spekulationsfrist).", "BMF 26.3.2004", None),
        ("Verkauf an Ehegatten/Familie nach 10 Jahren", '"ℹ Hinweis"',
         "Nach Ablauf der Frist steuerfreier Verkauf an Ehegatten → neue AfA-Basis beim Käufer; keine GrESt zwischen Ehegatten/Verwandten in gerader Linie. Muss fremdüblich finanziert sein (Gestaltungsmissbrauch-Risiko).", "§3 Nr. 4, 6 GrEStG; §42 AO", None),
        ("H  Eigennutzung", None, None, None, None),
        ("§35a Handwerkerleistungen", f'IF({{Modus}}="{EN}","✔ greift","– nur Eigennutzung (bei Vermietung: Mieter)")', "20 % der Lohnkosten, max. 1.200 €/Jahr; Zahlung nur per Überweisung.", "§35a Abs. 3 EStG", "INDEX([[Erm_35a]],2)"),
        ("§35a haushaltsnahe Dienstleistungen", f'IF({{Modus}}="{EN}","✔ greift","– nur Eigennutzung")', "Hausmeister, Treppenreinigung, Gartenpflege aus der Hausgeldabrechnung: 20 %, max. 4.000 €/Jahr.", "§35a Abs. 2 EStG", None),
        ("§35c energetische Sanierung", 'IF({Bonus35c_aktiv}=1,"✔ greift",IF({Modus}="Eigennutzung","⚠ prüfen (Gebäude > 10 J., keine KfW/BAFA-Förderung)","– nur Eigennutzung"))',
         "20 % über 3 Jahre (7/7/6 %), max. 40.000 € je Objekt; Bescheinigung Fachunternehmen; nicht mit Förderung kombinierbar. Maßnahmenabschluss bis Ende 2029.", "§35c EStG", None),
        ("Teilvermietung (Einliegerwohnung)", f'IF(AND({{Modus}}="{EN}",{{Verm_Anteil}}>0),"✔ greift","– nicht genutzt")',
         "Anteilige AfA, Zinsen und Kosten des vermieteten Teils sind Werbungskosten.", "§21 EStG", None),
        ("Häusliches Arbeitszimmer / Homeoffice", '"ℹ Hinweis"', "Jahrespauschale 1.260 € bei Mittelpunkt der Tätigkeit, sonst Tagespauschale 6 € (max. 1.260 €).", "§4 Abs. 5 Nr. 6b/6c EStG", None),
        ("Wohn-Riester", f'IF(AND({{Modus}}="{EN}",{{Riester}}>0),"✔ als EK genutzt","ℹ Hinweis")',
         "Riester-Kapital als Eigenkapital oder zur Tilgung; Zulagen 175 € + Kinderzulage. Nachgelagerte Besteuerung über das Wohnförderkonto.", "§92a EStG", None),
        ("I  Weitere Strukturen (nur Hinweis, nicht gerechnet)", None, None, None, None),
        ("Vermögensverwaltende GmbH / erweiterte Kürzung", '"ℹ Hinweis"',
         "Mieterträge in der GmbH: ca. 15,8 % KSt+Soli; Gewerbesteuer entfällt durch erweiterte Kürzung (§9 Nr. 1 S. 2 GewStG). Lohnt meist erst bei mehreren Objekten und Thesaurierung; Nachteil: keine steuerfreie Veräußerung nach 10 Jahren, Ausschüttung 25 % + Soli, Kosten, GrESt bei Einbringung.", "§9 Nr. 1 S. 2 GewStG", None),
        ("Schenkung/Erbschaft", '"ℹ Hinweis"', "Freibeträge alle 10 Jahre: Ehegatte 500.000 €, Kind 400.000 €, Enkel 200.000 €. Familienheim unter Bedingungen steuerfrei. Kettenschenkungen vermeiden.", "§§13, 16 ErbStG", None),
        ("Nießbrauch", '"ℹ Hinweis"', "Übertragung unter Vorbehaltsnießbrauch mindert den Schenkungswert; Eltern behalten Mieteinnahmen und AfA. Zuwendungsnießbrauch an Kinder kann Einkünfte auf niedrigere Steuersätze verlagern.", "§14 BewG; BMF 30.9.2013", None),
        ("Wann lohnt ein Steuerberater?", f'IF(OR({{Sonder7b_aktiv}}=1,{{Denkmal_BMG}}>0,{{Degr_aktiv}}=1,{{AN_ueber}}=1,AND({{Gewinn23_pflichtig}}=1,{{Verm_Anteil}}>0),{{Angehoerige}}="{JA}",{{Teilverm_Anteil}}>0,{{NB_an}}=1),"⚠ Empfehlung: ja (komplexe Punkte im Fall)","ℹ Erstjahr empfehlenswert")',
         "Sinnvoll bei: erstem Mietobjekt (Kaufpreisaufteilung!), Sonder-AfA, Denkmal, Teilvermietung, Angehörigen, Verkauf < 10 Jahre, mehreren Objekten, GmbH-Überlegungen. Kosten sind Werbungskosten.", "StBVV", None),
    ]
    for label, stat, expl, law, eff in items:
        if stat is None:
            m.section(ws, r, label, ncols=6)
            r += 1
            continue
        m.text(ws, r, 2, label, border=True, wrap=True, bold=True)
        c = m.view(ws, r, 3, stat, None, wrap=True)
        m.text(ws, r, 4, expl, border=True, wrap=True, size=9)
        m.text(ws, r, 5, law, border=True, wrap=True, size=9, color="404040")
        if eff:
            m.view(ws, r, 6, eff, F_EUR)
        else:
            m.text(ws, r, 6, "", border=True)
        ws.row_dimensions[r].height = 48
        r += 1
    rng = f"C5:C{r}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(C5,1)="✔"'], fill=fill(C_GREEN)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(C5,1)="✖"'], fill=fill(C_RED)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(C5,1)="⚠"'], fill=fill(C_YELLOW)))
    ws.freeze_panes = "C5"
    m.protect(ws)


def build_start(m):
    ws = m.wb.create_sheet("Start", 0)
    m.sheets["start"] = ws
    ws.sheet_view.showGridLines = False
    m.title(ws, "Immobilien-Rechner Deutschland 2026", "Kaufen oder nicht? – Kapitalanlage und Eigennutzung, inkl. Steuern, Szenarien und Vergleich mit ETF/Mieten.")
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 100
    r = 4
    m.section(ws, r, "Kurzanleitung", ncols=3); r += 1
    steps = [
        ("1. Eingaben", "Blatt 'Eingaben': gelbe Zellen mit deinen Werten überschreiben (die vorhandenen Zahlen sind ein Beispielfall). Oben 'Kapitalanlage' oder 'Eigennutzung' wählen."),
        ("2. Ergebnis", "Blatt 'Kennzahlen & Ampel': Fazit, Renditen, Cashflow, Vermögen, Diagramme und Warnhinweise."),
        ("3. Details", "'Kaufnebenkosten & Finanzierung' (Nebenkosten, Kaufpreisaufteilung, Tilgungsplan), 'Steuer' (AfA, Prüfungen, Verkauf), 'Cashflow-Plan' (alle Jahre), 'Kaufen vs. Mieten'."),
        ("4. Risiko", "'Szenarien & Sensitivität': Best/Base/Worst nebeneinander und zwei Raster (Zins × Kaufpreis, Miete × Wertsteigerung)."),
        ("5. Steuerhebel", "'Steuer-Checkliste': alle legalen Hebel mit Status für deinen Fall."),
        ("6. Gesetzeswerte", "'Parameter': Steuertarif, Grenzen, Sätze mit Quelle und Stand. Orange markierte Werte bitte selbst prüfen."),
        ("Blattschutz", "Alle Blätter sind ohne Passwort geschützt (Überprüfen → Blattschutz aufheben). Bitte nur bewusst aufheben."),
        ("Rechenlogik", "Jahr 1 = erste 12 Monate nach Übergabe; Darlehen mit Monatsraten exakt je Jahr; Steuer als Differenz ESt mit/ohne Objekt nach §32a (2026) inkl. Soli/KiSt; "
                        "Verkauf am Ende des Horizonts. Sensitivitäten rechnen das ganze Modell auf dem ausgeblendeten Blatt '_Sens' erneut."),
    ]
    for a, b in steps:
        m.text(ws, r, 2, a, bold=True, border=True)
        m.text(ws, r, 3, b, wrap=True, border=True)
        ws.row_dimensions[r].height = 32
        r += 1
    r += 1
    m.section(ws, r, "Farblegende", ncols=3); r += 1
    for txt, col, desc in (("1.234 €", C_IN, "Eingabe – darf geändert werden (entsperrt, mit Datenüberprüfung)"),
                           ("1.234 €", None, "Berechnung – gesperrt"),
                           ("1.234 €", C_KEY, "Wichtiges Ergebnis"),
                           ("PRÜFEN", C_ORANGE, "Parameter unsicher / Annahme – bitte verifizieren"),
                           ("LOHNT SICH", C_GREEN, "Ampel grün"), ("GRENZWERTIG", C_YELLOW, "Ampel gelb"),
                           ("LOHNT SICH NICHT", C_RED, "Ampel rot")):
        c = m.text(ws, r, 2, txt, border=True, fillc=col, color="0000FF" if col == C_IN else None, align="center")
        m.text(ws, r, 3, desc, border=True)
        r += 1
    r += 1
    m.section(ws, r, "Haftungshinweis", ncols=3); r += 1
    c = m.text(ws, r, 2, "Dieses Werkzeug ist eine vereinfachte Modellrechnung und KEINE Steuer-, Rechts- oder Anlageberatung. "
               "Gesetzliche Werte wurden nach bestem Wissen recherchiert (Stand 2026), können sich aber ändern oder im Einzelfall abweichen. "
               "Vereinfachungen u. a.: Modelljahre = Kalenderjahre, Kirchensteuer ohne Kappung/Kinderfreibeträge, Vorabpauschale beim ETF nicht gerechnet, "
               "Tilgungszuschuss vereinfacht, Kaufpreisaufteilung nur angelehnt an die BMF-Arbeitshilfe. Vor einer Kaufentscheidung Zahlen mit Steuerberater, "
               "Bank und ggf. Gutachter prüfen.", wrap=True, border=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.row_dimensions[r].height = 90
    r += 2
    m.text(ws, r, 2, "Erstellt", bold=True)
    m.text(ws, r, 3, f"{dt.date.today():%d.%m.%Y} – Generator: build_immobilienrechner.py")
    m.protect(ws)


# ====================================================================== main
def build(path, overrides=None):
    m = Model(overrides)
    build_parameter(m)
    build_eingaben(m)
    build_finanzierung(m)  # benötigt Engine-Namen nur als Z_-Namen -> werden später angelegt
    build_steuer_scalars(m)
    build_szen_scalars(m)
    build_kosten_scalars(m)
    m.engine = engine_columns()
    # Engine-Namen vorab registrieren (für Prüfungen)
    for c in m.engine:
        m.nodes["Z_" + c["key"]] = {"tpl": None, "kind": "engine"}
    build_cashflow(m)
    build_steuer_post(m)
    build_vergleich(m)
    build_dashboard(m)
    build_sensitivity(m)
    build_checkliste(m)
    build_start(m)
    order = ["Start", "Eingaben", "Kaufnebenkosten & Finanzierung", "Steuer", "Cashflow-Plan", "Kaufen vs. Mieten",
             "Kennzahlen & Ampel", "Szenarien & Sensitivität", "Steuer-Checkliste", "Parameter", "_Sens"]
    missing = sorted(n for n in m.used if n not in m.nodes)
    assert not missing, f"Unbekannte Namen: {missing}"
    m.wb._sheets = [m.wb[n] for n in order]
    m.wb.active = 0
    m.wb.calculation.fullCalcOnLoad = True
    m.wb.calculation.calcMode = "auto"
    m.wb.save(path)
    return m


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Immobilienrechner_2026.xlsx"
    ov = {}
    for a in sys.argv[2:]:
        k, v = a.split("=", 1)
        try:
            v = float(v) if "." in v else int(v)
        except ValueError:
            pass
        ov[k] = v
    build(out, ov)
    print("geschrieben:", out)
