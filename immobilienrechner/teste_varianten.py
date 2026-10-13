#!/usr/bin/env python3
"""Baut Varianten, rechnet sie mit LibreOffice neu und prüft auf Fehler + Plausibilität."""
import datetime as dt
import os
import shutil
import subprocess
import sys

import build_immobilienrechner as B
import test_rechner as T
from pruefe_rechner import load

RECALC = "/root/.claude/skills/synced/80589829-e181-4b8f-8d3b-74fa4a5ef02c_9c8e0ddd-49e5-487f-8836-8bf3f0d3ab68/xlsx/scripts/recalc.py"
OUT = sys.argv[1] if len(sys.argv) > 1 else "/tmp/varianten"
os.makedirs(OUT, exist_ok=True)

VARIANTEN = {
    "01_basis": {},
    "02_eigennutzung": {"Modus": "Eigennutzung"},
    "03_eigen_teilverm": {"Modus": "Eigennutzung", "Teilverm_Anteil": 0.3, "Teilverm_Miete": 450, "Teilverm_Zuordnung": "Ja",
                          "Renovierung": 30000, "Renov_Energ": 0.5, "Objekttyp": "Haus (EFH/DHH/RH)", "Grundstueck": 450},
    "04_neubau_degr_7b": {"Baujahr": 2026, "Baubeginn": dt.date(2024, 5, 1), "Erwerb_FJ": "Ja", "Degr_Wahl": "Ja",
                          "Sonder7b_Wahl": "Ja", "Bauantrag": dt.date(2024, 1, 15), "Effizienz_EH40": "Ja",
                          "Kaufpreis": 420000, "Zustand": "Neubau / neuwertig", "Renovierung": 0},
    "05_denkmal": {"Denkmal_Status": "Baudenkmal (§7i / §10f)", "Denkmal_Betrag": 150000, "Baujahr": 1900},
    "06_nachbeleihung_etf": {"NB_Aktiv": "Ja", "NB_Verwendung": "ETF/Wertpapiere", "Horizont": 15},
    "07_nachbeleihung_mod": {"NB_Aktiv": "Ja", "NB_Verwendung": "Modernisierung dieses Objekts", "NB_Mod_Art": "Herstellungsaufwand",
                             "NB_Modus": "Fester Betrag", "NB_Betrag": 40000},
    "08_angehoerige": {"Angehoerige": "Ja", "Kaltmiete": 500},
    "09_worst_kurz": {"Szenario": "Worst", "Horizont": 8, "RND_Gutachten": "Ja", "RND_Jahre": 30, "Veranlagung": "Zusammenveranlagung",
                      "Kirchensteuer": "Ja", "zvE": 140000},
    "10_direkt_anschaffungsnah": {"Steuer_Modus": "Grenzsteuersatz direkt", "Renovierung": 40000, "Verteilung_82b": 3,
                                  "Erh_Jahr": 2, "Erh_Betrag": 5000},
    "11_kfw_disagio_sonder": {"Kredit4_Betrag": 100000, "Kredit4_TZ": 0.15, "Kredit4_TF": 2, "Kredit3_Betrag": 50000,
                              "Kredit3_Disagio": 0.05, "Kredit1_ST": 5000, "Kredit1_ST_ab": 2, "Kredit1_ST_bis": 5,
                              "Kredit2_BZ_Monate": 6},
    "12_extrem": {"Eigenkapital": 2000000, "Horizont": 1, "Kaltmiete": 0, "Kredit2_Betrag": 0, "Wertsteigerung": -0.05,
                  "Leerstand": 1, "zvE": 0},
}


def run(name, ov):
    p = os.path.join(OUT, f"{name}.xlsx")
    B.build(p, ov)
    res = subprocess.run([sys.executable, RECALC, p, "300"], capture_output=True, text=True)
    if '"status": "success"' not in res.stdout:
        print(name, "RECALC:", res.stdout[:800], res.stderr[-800:])
        return False
    wb, N, errs = load(p)
    print(f"\n=== {name} === Fehlerzellen: {len(errs)} {errs[:5]}")
    ok = T.main(p) if not errs else False
    for k in ("Modus", "Szenario", "AfA_Art", "WK_Quote", "W_Verbilligt", "W_AN", "W_7b", "W_Degr", "Steuer_1", "CF_nach_1_mtl",
              "IRR_Wert", "Steuer_Verkauf", "EndA", "EndB", "Vorteil", "Fazit", "Szen_Base_IRR_Wert", "Sens_A_3_3_IRR_Wert",
              "Sens_B_3_3_Vorteil"):
        print(f"   {k:22s} {N.get(k)}")
    # Schattenrechnung Mitte == Hauptmodell (nur wenn aktives Szenario Base)
    if N["Szenario"] == "Base":
        for o in ("IRR_Wert", "Vorteil", "CF_nach_1_mtl", "Vermoegen_N"):
            for pre in ("Szen_Base", "Sens_A_3_3", "Sens_B_3_3"):
                a, b = N[f"{pre}_{o}"], N[o]
                same = (a == b) if isinstance(a, str) or isinstance(b, str) else abs(a - b) < 1e-6
                if not same:
                    print("   FEHLER Schatten ungleich:", pre, o, a, b)
                    ok = False
    return ok and not errs


if __name__ == "__main__":
    sel = sys.argv[2:] or list(VARIANTEN)
    results = {n: run(n, VARIANTEN[n]) for n in sel}
    print("\nZUSAMMENFASSUNG:", results)
    sys.exit(0 if all(results.values()) else 1)
