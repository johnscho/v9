# Immobilien-Rechner Deutschland 2026

`Immobilienrechner_2026.xlsx` – reines .xlsx (keine Makros) für Excel 365 Desktop und Excel Online.
Entscheidungshilfe für Kauf als **Kapitalanlage** oder zur **Eigennutzung** (umschaltbar auf dem Blatt *Eingaben*).

**Bedienung:** Nur die gelben Zellen ändern (Blatt *Eingaben*, Best/Worst und Rasterachsen auf *Szenarien & Sensitivität*,
gesetzliche Werte auf *Parameter*). Ergebnis auf *Kennzahlen & Ampel*. Blattschutz ohne Passwort.
Beim ersten Öffnen einer heruntergeladenen Datei „Bearbeitung aktivieren“ klicken, damit Excel neu rechnet.

Keine Steuerberatung – siehe Haftungshinweis auf dem Blatt *Start*.

## Dateien

| Datei | Zweck |
|---|---|
| `build_immobilienrechner.py` | Generator (openpyxl). Alle Formeln, Namen, Prüfungen, Schutz. |
| `erstelle_final.py` | Baut, rechnet mit LibreOffice headless, prüft auf Fehlerwerte, schreibt Ergebniswerte als Cache in die Datei. |
| `pruefe_rechner.py` | Liest eine berechnete Mappe, meldet Fehlerzellen, gibt Kernergebnisse aus. |
| `test_rechner.py` | Unabhängige Python-Nachrechnung (Tarif §32a, Steuerwirkung, Annuitäten, Cashflow, IRR). |
| `teste_varianten.py` | 12 Varianten (Eigennutzung, Neubau degressiv + §7b, Denkmal, Nachbeleihung, Angehörige, Worst …). |

```bash
python3 erstelle_final.py Immobilienrechner_2026.xlsx   # neu erzeugen
python3 teste_varianten.py /tmp/varianten               # Testsuite
```
