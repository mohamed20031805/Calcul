"""
TEST — Vérification des statuts par Clé Unique
================================================
Lance ce script après hits_traitement.py pour valider
que la logique des statuts a bien été appliquée.

Usage :
    python test_statuts.py
    python test_statuts.py output/HITS_rapport_Q2-2025.xlsx   (chemin custom)
"""

import sys
import os
import pandas as pd

# ─── Config ───────────────────────────────────────────────
DEFAULT_FILE       = "output/HITS_rapport_Q1-2025.xlsx"
SHEET_DONNEES      = "Données"
STATUS_OPEN        = "To be treated"
DATE_FIN_TRIMESTRE = "2025-03-31"        # doit correspondre à hits_traitement.py
# ──────────────────────────────────────────────────────────

GREEN  = "\033[92m"
YELLOW = "\033[93m"
RED    = "\033[91m"
CYAN   = "\033[96m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def ok(msg):    print(f"  {GREEN}✓{RESET} {msg}")
def warn(msg):  print(f"  {YELLOW}⚠{RESET} {msg}")
def err(msg):   print(f"  {RED}✗{RESET} {msg}")
def info(msg):  print(f"  {CYAN}→{RESET} {msg}")


def find_col(df, hint):
    for col in df.columns:
        if hint.lower() in str(col).lower():
            return col
    return None


def run_tests(filepath: str):
    print()
    print(BOLD + "=" * 62 + RESET)
    print(BOLD + "  TEST — Vérification statuts par Clé Unique" + RESET)
    print(BOLD + "=" * 62 + RESET)
    print(f"\n  Fichier : {filepath}")

    # ── Chargement ────────────────────────────────────────
    if not os.path.exists(filepath):
        err(f"Fichier introuvable : {filepath}")
        sys.exit(1)

    df = pd.read_excel(filepath, sheet_name=SHEET_DONNEES)
    print(f"  Lignes chargées : {len(df)}")

    col_key = find_col(df, "Clé Unique")
    col_st  = find_col(df, "Status")
    col_ud  = find_col(df, "Update Date")
    col_ig  = find_col(df, "Integration Date")

    if not col_key:
        err("Colonne 'Clé Unique' introuvable — vérifie le fichier.")
        sys.exit(1)
    if not col_st:
        err("Colonne 'Status' introuvable — vérifie le fichier.")
        sys.exit(1)

    # Parser dates
    df[col_ud] = pd.to_datetime(df[col_ud], errors="coerce")
    df[col_ig] = pd.to_datetime(df[col_ig], errors="coerce")
    date_fin   = pd.Timestamp(DATE_FIN_TRIMESTRE)

    # ── Comptage statuts par clé ──────────────────────────
    counts      = df.groupby(col_key)[col_st].count()
    cles_1      = counts[counts == 1]
    cles_2      = counts[counts == 2]
    cles_3plus  = counts[counts >= 3]

    total_cles  = len(counts)

    print()
    print(BOLD + "  [1] Répartition des clés par nombre de statuts" + RESET)
    print(f"      Total clés uniques   : {total_cles}")
    print(f"      Clés avec 2 statuts  : {len(cles_2)}")
    print(f"      Clés avec 1 statut   : {len(cles_1)}")
    print(f"      Clés avec 3+ statuts : {len(cles_3plus)}")

    # ── TEST A — Aucune clé à 3+ statuts ─────────────────
    print()
    print(BOLD + "  [TEST A] Aucune clé ne doit avoir 3+ statuts" + RESET)
    if len(cles_3plus) == 0:
        ok(f"Aucune clé avec 3+ statuts — suppression rang 3+ correcte")
    else:
        err(f"{len(cles_3plus)} clé(s) avec 3+ statuts détectée(s) — la suppression n'a pas fonctionné !")
        for key, cnt in cles_3plus.items():
            rows = df[df[col_key] == key][[col_st, col_ud]].to_string(index=False)
            print(f"\n      Clé : {str(key)[:70]}")
            print(f"      {rows}")

    # ── TEST B — Clés à 1 statut = dossiers ouverts légitimes ────────
    print()
    print(BOLD + "  [TEST B] Clés à 1 statut — vérification 'To be treated' + date fin trimestre" + RESET)

    erreurs_b = []
    for key in cles_1.index:
        row = df[df[col_key] == key].iloc[0]
        statut = str(row[col_st]).strip()
        ud     = row[col_ud]

        # Le statut doit être STATUS_OPEN
        if statut != STATUS_OPEN:
            erreurs_b.append((key, f"statut inattendu '{statut}' (attendu '{STATUS_OPEN}')"))
            continue

        # Update Date doit être remplacée par date_fin
        if pd.isna(ud):
            erreurs_b.append((key, "Update Date est vide (date fin trimestre non appliquée)"))
        elif ud.normalize() != date_fin.normalize():
            erreurs_b.append((key, f"Update Date = {ud.date()} (attendu {date_fin.date()})"))

    if not erreurs_b:
        if len(cles_1) == 0:
            ok("Aucun dossier ouvert (toutes les clés ont 2 statuts)")
        else:
            ok(f"{len(cles_1)} dossier(s) ouvert(s) — statut '{STATUS_OPEN}' et date fin {date_fin.date()} corrects")
    else:
        for key, msg in erreurs_b:
            err(f"Clé : {str(key)[:60]}... → {msg}")

    # ── TEST C — Clés à 2 statuts : statut 1 = To be treated ─────────
    print()
    print(BOLD + "  [TEST C] Clés à 2 statuts — statut 1 doit être 'To be treated'" + RESET)

    erreurs_c = []
    for key in cles_2.index:
        rows_key = df[df[col_key] == key].sort_values(col_ud)
        st1 = str(rows_key.iloc[0][col_st]).strip()
        st2 = str(rows_key.iloc[1][col_st]).strip()

        if st1 != STATUS_OPEN:
            erreurs_c.append((key, f"statut 1 = '{st1}' (attendu '{STATUS_OPEN}')"))
        if not st2 or st2 in ("nan", "None", ""):
            erreurs_c.append((key, f"statut 2 est vide"))

    if not erreurs_c:
        ok(f"{len(cles_2)} clé(s) à 2 statuts — ordre correct")
    else:
        for key, msg in erreurs_c:
            err(f"Clé : {str(key)[:60]}... → {msg}")

    # ── TEST D — diff_day cohérence ───────────────────────
    print()
    print(BOLD + "  [TEST D] diff_day — cohérence avec les dates" + RESET)

    col_dd = find_col(df, "diff_day")
    if not col_dd:
        warn("Colonne diff_day introuvable — test ignoré")
    else:
        df["_dd_calc"] = (df[col_ud].dt.normalize() - df[col_ig].dt.normalize()).dt.days
        ecarts = df[df["_dd_calc"] != df[col_dd]]
        if len(ecarts) == 0:
            ok("diff_day cohérent sur toutes les lignes")
        else:
            err(f"{len(ecarts)} ligne(s) avec diff_day incohérent :")
            print(ecarts[[col_key, col_ig, col_ud, col_dd, "_dd_calc"]].head(5).to_string(index=False))
        df.drop(columns=["_dd_calc"], inplace=True)

    # ── TEST E — doublons de clé à même rang ─────────────
    print()
    print(BOLD + "  [TEST E] Pas de doublon de (Clé Unique + Statut + Update Date)" + RESET)
    dupes = df.duplicated(subset=[col_key, col_st, col_ud])
    if dupes.sum() == 0:
        ok("Aucun doublon détecté")
    else:
        err(f"{dupes.sum()} doublon(s) détecté(s) :")
        print(df[dupes][[col_key, col_st, col_ud]].head(5).to_string(index=False))

    # ── RÉSUMÉ ────────────────────────────────────────────
    all_errors = len(cles_3plus) + len(erreurs_b) + len(erreurs_c)
    print()
    print(BOLD + "=" * 62 + RESET)
    if all_errors == 0:
        print(BOLD + GREEN + "  RÉSULTAT GLOBAL : ✅  TOUS LES TESTS SONT PASSÉS" + RESET)
    else:
        print(BOLD + RED + f"  RÉSULTAT GLOBAL : ❌  {all_errors} ANOMALIE(S) DÉTECTÉE(S)" + RESET)
    print(BOLD + "=" * 62 + RESET)
    print()

    return all_errors


if __name__ == "__main__":
    filepath = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FILE
    errors = run_tests(filepath)
    sys.exit(0 if errors == 0 else 1)
