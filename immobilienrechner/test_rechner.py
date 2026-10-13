#!/usr/bin/env python3
"""Unabhängige Nachrechnung (Python) ausgewählter Größen gegen die mit LibreOffice berechnete Mappe."""
import math
import sys

from pruefe_rechner import load


def tarif(x):  # §32a EStG 2026
    x = math.floor(x)
    if x <= 12348:
        return 0
    if x <= 17799:
        y = (x - 12348) / 1e4
        return math.floor((914.51 * y + 1400) * y)
    if x <= 69878:
        z = (x - 17799) / 1e4
        return math.floor((173.10 * z + 2397) * z + 1034.87)
    if x <= 277825:
        return math.floor(0.42 * x - 11135.63)
    return math.floor(0.45 * x - 19470.38)


def est(x, split=1):
    return split * tarif(x / split)


def soli(e, split=1):
    fg = 20350 * split
    return 0 if e <= fg else min(0.055 * e, 0.119 * (e - fg))


def annuity_year(rs, i, ann):
    """Monatliche Raten über 12 Monate."""
    zins = 0
    for _ in range(12):
        z = rs * i / 12
        t = min(ann / 12 - z, rs)
        zins += z
        rs -= t
    return rs, zins


def main(path):
    wb, N, errs = load(path)
    ok = True

    def check(label, got, exp, tol=1.0):
        nonlocal ok
        good = abs(got - exp) <= tol
        ok &= good
        print(f"{'OK ' if good else 'FEHLER'} {label:45s} Mappe={got:14.2f}  Python={exp:14.2f}")

    assert not errs, errs[:10]
    split = N["Split"]
    kist = N["KiSt_Satz_eff"]
    for x in (10000, 15000, 40000, 70000, 150000, 300000):
        print(f"   Tarif-Stichprobe zvE {x}: {est(x, 1)} € (einzeln) / {est(x, 2)} € (Splitting)")
    check("ESt auf zvE", N["ESt_zvE"], est(N["zvE"], split))
    e = est(N["zvE"], split)
    check("ESt+Soli+KiSt auf zvE", N["Steuer_zvE"], e + soli(e, split) + e * kist, 0.05)
    # Jahr 1 Steuerwirkung
    zv = N["Z_zvE_mit"][1]
    e1 = max(0, est(zv, split) - N["Z_Erm_35a"][1] - N["Z_Erm_35c"][1])
    e0 = est(N["Z_zvE_ohne"][1], split)
    sw = (e1 + soli(e1, split) + e1 * kist) - (e0 + soli(e0, split) + e0 * kist)
    if N["Steuer_Modus"].startswith("Aus zvE"):
        check("Steuerwirkung Jahr 1", N["Z_Steuerwirkung"][1], sw, 0.05)
    # Darlehen 1: Jahr 1..3 nachrechnen (ohne Sondertilgung/TF)
    rs = N["Kredit1_Betrag"]
    i = N["Z_L1_Zins"][1]
    ann = rs * (i + N["Kredit1_Tilgung"])
    for t in (1, 2, 3):
        rs, z = annuity_year(rs, i, ann)
        check(f"Darlehen 1 Restschuld Ende Jahr {t}", N["Z_L1_RS_E"][t] + N["Z_L1_Sonder"][t], rs, 1.0)
        check(f"Darlehen 1 Zinsen Jahr {t}", N["Z_L1_Zinsen"][t], z, 1.0)
        rs -= N["Z_L1_Sonder"][t]
    # Konsistenz: Restschuld = Summe Einzeldarlehen; Restschuld nie negativ
    for t in range(41):
        s = sum(N[f"Z_L{k}_RS_E"][t] for k in range(1, 6))
        assert abs(s - N["Z_RS_Summe"][t]) < 0.01
        assert min(N[f"Z_L{k}_RS_E"][t] for k in range(1, 6)) >= -0.01
    print("OK  Restschulden konsistent und nie negativ")
    # Tilgungsplan: Anfangsschuld - Summe Tilgungen - Zuschüsse = Endschuld (inkl. Nachbeleihung)
    t = 40
    lhs = N["Z_RS_Summe"][0] + sum(N["Z_NB_Ausz"][1:]) - sum(N["Z_Tilgung_Summe"][1:]) - sum(N["Z_Zuschuss_Summe"][1:])
    check("Tilgungsplan summiert sich auf", N["Z_RS_Summe"][t], lhs, 1.0)
    # Cashflow-Identität
    t = 1
    cf = (N["Z_Miete_Ist"][t] - N["Z_U_Traeger"][t] - N["Z_N_Kosten"][t] - N["Z_Ruecklage"][t] - N["Z_Instand"][t]
          - N["Z_Erh_gross"][t] - N["Z_S_Kosten"][t] - N["Z_Bereit"][t] - N["Z_Zinsen_Summe"][t] - N["Z_Tilgung_Summe"][t])
    check("Cashflow vor Steuern Jahr 1", N["Z_CF_vor"][1], cf, 0.01)
    # IRR nachrechnen
    flows = N["Z_IRR_Flow"]
    if isinstance(N["IRR_Wert"], (int, float)):
        r = N["IRR_Wert"]
        npv = sum(f / (1 + r) ** k for k, f in enumerate(flows))
        check("NPV bei IRR ≈ 0", npv, 0, 1.0)
    # AfA Summe <= Bemessungsgrundlage
    assert sum(N["Z_Geb_AfA"][1:]) + sum(N["Z_S7b"][1:]) <= max(N["Z_Geb_BMG"][1:]) + 1
    print("OK  Gebäude-AfA übersteigt nie die Bemessungsgrundlage")
    print("GESAMT:", "BESTANDEN" if ok else "FEHLER")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1]) else 1)
