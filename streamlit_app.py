import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from collections import defaultdict
import datetime, io

st.set_page_config(page_title="QC Monitor", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.block-container{padding-top:1.5rem}
div[data-testid="metric-container"]{background:#0f1117;border:1px solid #1c2030;border-radius:8px;padding:12px 16px}
div[data-testid="metric-container"] label{font-size:11px!important;color:#7c879e!important;text-transform:uppercase;letter-spacing:.5px}
</style>""", unsafe_allow_html=True)

# ─── HELPERS ─────────────────────────────────────────────────────────────────

SKIP_ROWS = {'Total','Total Panel Dédupliqué','Total Panel Dédupliqué - Top 5 Sites',
    'Total Panel Dédupliqué  - Top 11 Sites','Total Panel Dédupliqué Marché',
    'Immobilier Notaire','Immonot','Site','Département','Totaux','Total Panel Dedup'}

SITES = ['AvendreAlouer',"Bien'ici",'Figaro Immo','Green-Acres','Leboncoin',
         'LogicImmo','MeilleursAgents','OuestFrance','PAP','ParuVendu','SeLoger','SuperImmo']

def excel_date_str(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, str): return v.strip()
    if isinstance(v,(int,float)) and 40000<v<50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def get_sheet_data(ws, site_col=2, start_col=3):
    """Extract site data + months from a sheet."""
    header_row = None
    for r in range(1, min(20, ws.max_row+1)):
        v = ws.cell(r, site_col).value
        if v in ('Site','Département','Région'):
            header_row = r; break
    if not header_row: return None, None, None

    month_cols, months = [], []
    for c in range(start_col, ws.max_column+1):
        v = ws.cell(header_row, c).value
        if v is None: continue
        if (isinstance(v, datetime.datetime) or
            (isinstance(v,(int,float)) and 40000<v<50000) or
            (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
            month_cols.append(c); months.append(excel_date_str(v))
    if len(months)<2: return None, None, None

    data = {}
    for r in range(header_row+1, ws.max_row+1):
        site = ws.cell(r, site_col).value
        if not site or not isinstance(site,str): continue
        site = site.strip()
        if not site: continue
        vm = ws.cell(r, month_cols[-1]).value
        vm1 = ws.cell(r, month_cols[-2]).value
        if isinstance(vm,(int,float)) and isinstance(vm1,(int,float)):
            data[site] = {'m':float(vm),'m1':float(vm1)}
    return data, months[-1], months[-2]

def check_ok(name, condition, detail, severity="error", sheet=""):
    return {'check':name,'status':'ok' if condition else 'fail',
            'severity':severity if not condition else 'ok','detail':detail,'sheet':sheet}

# ─── QC GOLD CHECKS ──────────────────────────────────────────────────────────



"""
QC Gold V2 — Exact replicas of Panel Checker controls applied to HAM files.
"""
import io, datetime
from openpyxl import load_workbook

def excel_date_str(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, str): return v.strip()
    if isinstance(v,(int,float)) and 40000<v<50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def get_val_by_label(ws, label, col_offset=0, search_cols=range(1,5)):
    """Find row by label in any of search_cols, return value from last data col + offset."""
    for r in range(1, ws.max_row+1):
        for c in search_cols:
            cell_val = ws.cell(r,c).value
            if cell_val and str(cell_val).strip() == label:
                # Find last non-None column
                last_c = ws.max_column
                while last_c > c and ws.cell(r, last_c).value is None:
                    last_c -= 1
                target_c = last_c + col_offset
                if target_c > 0:
                    return ws.cell(r, target_c).value
    return None

def get_val_contains(ws, label_part, col_offset=0, search_cols=range(1,5)):
    """Find row where label contains label_part."""
    for r in range(1, ws.max_row+1):
        for c in search_cols:
            cell_val = ws.cell(r,c).value
            if cell_val and label_part.lower() in str(cell_val).lower():
                last_c = ws.max_column
                while last_c > c and ws.cell(r, last_c).value is None:
                    last_c -= 1
                target_c = last_c + col_offset
                if target_c > 0:
                    return ws.cell(r, target_c).value, str(cell_val).strip()
    return None, None

def get_last_two_months(ws, site_col=2):
    """Get last and second-to-last month columns and their labels."""
    header_row = None
    for r in range(1, min(20, ws.max_row+1)):
        v = ws.cell(r, site_col).value
        if v in ('Site','Département','Région'):
            header_row = r; break
    if not header_row: return None, None, None, None

    month_cols, months = [], []
    for c in range(site_col+1, ws.max_column+1):
        v = ws.cell(header_row, c).value
        if v is None: continue
        if (isinstance(v, datetime.datetime) or
            (isinstance(v,(int,float)) and 40000<v<50000) or
            (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
            month_cols.append(c); months.append(excel_date_str(v))
    if len(months)<2: return None, None, None, None
    return month_cols[-1], months[-1], month_cols[-2], months[-2]

def get_site_val(ws, site_name, site_col=2):
    """Get last month value for a site."""
    lc, lm, pc, pm = get_last_two_months(ws, site_col)
    if not lc: return None, None, None, None
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, site_col).value
        if v and str(v).strip() == site_name:
            return ws.cell(r,lc).value, ws.cell(r,pc).value, lm, pm
    return None, None, lm, pm

def check(name, ok, detail, severity="error", sheet=""):
    return {'check':name, 'status':'ok' if ok else 'fail',
            'severity':severity if not ok else 'ok', 'detail':detail, 'sheet':sheet}

def is_close(a, b, tol=0.01):
    if a is None or b is None: return False
    if a == 0 and b == 0: return True
    if a == 0 or b == 0: return False
    return abs(a-b)/max(abs(a),abs(b)) < tol

# ─── HELPERS FOR SOURCE-FILE CHECKS ─────────────────────────────────────────

def get_lc(ws):
    """Son ay kolonunu (lc) ve etiketini (lm) döndür."""
    for r in range(1, 15):
        mc = []
        for c in range(1, min(ws.max_column+1, 200)):
            v = ws.cell(r, c).value
            if v is None: continue
            if (isinstance(v, datetime.datetime) or
                (isinstance(v,(int,float)) and 40000<v<50000) or
                (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
                mc.append((c, excel_date_str(v)))
        if len(mc) >= 3:
            return mc[-1][0], mc[-1][1]
    return None, None

def row_val(ws, label, lc, col=2, skip=0):
    """col kolonunda tam eşleşen etiketi bul, lc kolonundaki değeri döndür."""
    found = 0
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, col).value
        if v and isinstance(v, str) and v.strip() == label:
            if found == skip:
                return ws.cell(r, lc).value
            found += 1
    return None

def row_val_contains(ws, parts, lc, col=2):
    """Etikette 'parts' kelimelerini içeren ilk satırın lc değerini döndür."""
    parts = parts if isinstance(parts, list) else [parts]
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, col).value
        if v and isinstance(v, str):
            if all(p.lower() in v.lower() for p in parts):
                return ws.cell(r, lc).value
    return None

def fmt(v):
    if v is None: return 'N/A'
    try: return f"{float(v):,.0f}"
    except: return str(v)

# ─── TAB 3.1 CHECKS ──────────────────────────────────────────────────────────
# Source des formules Panel Checker (vérifiées sur le fichier réel):
#
# tab 3.1 row 21 "contrôle type":   D21=D8+D9+D10+D11=D12
#   → 3.1.1: Total identifiés Joreca = Agence + Intermédiaire + Notaire + Autres (Annonces col)
#
# tab 3.1 row 22 "contrôle identification":  D22=D18=D15+D16+D17
#   → 3.1.1: Total général pros = Pros identifiés + Pros à identifier + Annonces incomplètes
#
# tab 3.1 row 23 "contrôle tab 1":  D23: D18='tab 1'!O116 (Bien'ici)
#                                   AQ23: AQ18='tab 1'!O130 (Total Dédup Marché)
#   → 3.1.1 Total général professionnels (Annonces col) per site
#   == 1.2 Pro_Part Ancien annonces per site (last month col)

def _311_site_cols(ws311):
    """3.1.1 Pros par site: retourne dict site→(Pro_col, Annonces_col, Moyenne_col)"""
    cols = {}
    for c in range(2, ws311.max_column+1):
        v = ws311.cell(1, c).value
        if v and isinstance(v, str) and v.strip():
            name = v.strip().strip("'\"")
            h2 = ws311.cell(2, c).value
            if h2 and str(h2).strip() == 'Pro':
                cols[name] = (c, c+1, c+2)
    return cols

def _311_row(ws311, label):
    """3.1.1: retourne numéro de ligne dont col1 == label"""
    for r in range(1, ws311.max_row+1):
        v = ws311.cell(r, 1).value
        if v and str(v).strip() == label:
            return r
    return None

def check_tab31(wb31_bytes, wb1_bytes):
    checks = []
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)
    wb1  = load_workbook(io.BytesIO(wb1_bytes), data_only=True)

    ws311 = wb31['3.1.1 Pros par site ']
    ws314 = wb31['3.1.4 Evolution Pros par type']
    ws12  = wb1['1.2 Pro_Part']

    lc12, lm = get_lc(ws12)
    if not lc12: return checks

    site_cols = _311_site_cols(ws311)

    SITES = ['AvendreAlouer', "Bien'ici", 'Figaro Immo', 'Green-Acres', 'Leboncoin',
             'LogicImmo', 'MeilleursAgents', 'OuestFrance', 'PAP', 'ParuVendu',
             'SeLoger', 'Superimmo']
    # 1.2 utilise 'SuperImmo' (majuscule I), 3.1.1 utilise 'Superimmo'
    ALIAS_12 = {'Superimmo': 'SuperImmo'}

    # ── CHECK A: contrôle tab 1 (row 23) — site par site
    # D23: D18 (Bien'ici Total général annonces dans 3.1.1) = 'tab 1'!O116 (1.2 Bien'ici annonces)
    # Vérifie que chaque site dans 3.1.1 "Total général professionnels" Annonces col
    # correspond bien à la même valeur dans 1.2 Pro_Part Ancien

    r_total_gen = _311_row(ws311, 'Total général professionnels')
    if r_total_gen:
        for site in SITES:
            if site not in site_cols: continue
            ann_col = site_cols[site][1]
            v311 = ws311.cell(r_total_gen, ann_col).value
            site_12 = ALIAS_12.get(site, site)
            v12 = row_val(ws12, site_12, lc12)
            if v311 is not None and v12 is not None:
                ok = (float(v311) == float(v12))
                checks.append(check(
                    f"[row 23] {site} — 3.1.1 Total général = 1.2 annonces {lm}",
                    ok,
                    (f"✅ {site} cohérent : {fmt(v311)} annonces." if ok else
                     f"❌ 3.1.1 'Total général professionnels' Annonces = {fmt(v311)}"
                     f" | 1.2 '{site}' = {fmt(v12)}"
                     f" | Différence = {fmt(abs(float(v311)-float(v12)))} annonces"),
                    "error", "tab 3.1 — contrôle tab 1"
                ))

    # ── CHECK A bis: AQ23 — Total Panel Dédupliqué Marché
    # AQ23: AQ18 = 'tab 1'!O130
    # tab 3.1 AQ18 = 3.1.1 col 42 header="Annonces dédoublonnées", row 14 "Total général professionnels"
    # 'tab 1'!O130 = 1.2 Pro_Part "Total Panel Dédupliqué Marché" dernière col
    #
    # ATTENTION: 3.1.1 a deux colonnes contenant "dédoublonn":
    #   col 41 = "Pros dédoublonnés"   → row 14 = 89,056  (nombre d'annonceurs)
    #   col 42 = "Annonces dédoublonnées" → row 14 = 1,406,035 (nombre d'annonces) ← C'EST CELUI-CI
    # On cherche la colonne dont row 2 = "Annonces dédoublonnées" (pas "Pros dédoublonnés")
    dedup_col = None
    for c in range(1, ws311.max_column+1):
        h = ws311.cell(2, c).value
        if h and isinstance(h, str) and h.strip().lower() == 'annonces dédoublonnées':
            dedup_col = c; break

    if dedup_col and r_total_gen:
        v311_dedup = ws311.cell(r_total_gen, dedup_col).value
        v12_dedup  = row_val(ws12, 'Total Panel Dédupliqué Marché', lc12)
        if v311_dedup is not None and v12_dedup is not None:
            ok = (float(v311_dedup) == float(v12_dedup))
            checks.append(check(
                f"[AQ23] 3.1.1 Annonces dédoublonnées ({fmt(v311_dedup)}) = 1.2 Total Panel Dédupliqué Marché ({fmt(v12_dedup)}) — {lm}",
                ok,
                (f"✅ Les deux fichiers sont cohérents : {fmt(v311_dedup)} annonces dédoublonnées." if ok else
                 f"❌ 3.1.1 col 'Annonces dédoublonnées' (Total général) = {fmt(v311_dedup)}"
                 f" | 1.2 'Total Panel Dédupliqué Marché' = {fmt(v12_dedup)}"
                 f" | Différence = {fmt(abs(float(v311_dedup)-float(v12_dedup)))} annonces"
                 f" → Demander à l'équipe de vérifier le fichier 3.1 ou 1."),
                "error", "tab 3.1 — contrôle tab 1 (AQ23)"
            ))

    # ── CHECK B: contrôle type (row 21)
    # D21=D8+D9+D10+D11=D12
    # 3.1.1: Total identifiés Joreca = Agence + Intermédiaire + Notaire + Autres (Annonces col)
    r_total_id  = _311_row(ws311, 'Total identifiés base Joreca')
    r_agence    = _311_row(ws311, 'Agence immobilière')
    r_interm    = _311_row(ws311, 'Intermédiaire')
    r_notaire   = _311_row(ws311, 'Notaire')
    r_autres    = _311_row(ws311, 'autres (uniquement le Autres Hors Promoteur et Constructeur)')
    if r_autres is None:
        r_autres = _311_row(ws311, 'Autres')

    if all(r is not None for r in [r_total_id, r_agence, r_interm, r_notaire, r_autres]):
        for site in SITES:
            if site not in site_cols: continue
            ann_col = site_cols[site][1]
            t  = ws311.cell(r_total_id, ann_col).value
            ag = ws311.cell(r_agence,   ann_col).value
            it = ws311.cell(r_interm,   ann_col).value
            no = ws311.cell(r_notaire,  ann_col).value
            au = ws311.cell(r_autres,   ann_col).value
            vals = [t, ag, it, no, au]
            if all(isinstance(v, (int,float)) for v in vals) and float(t) > 0:
                soma = float(ag) + float(it) + float(no) + float(au)
                ok   = (soma == float(t))
                checks.append(check(
                    f"[tab 1 row 21] {site}: Agences+Interm+Notaires+Autres = Total identifiés ({lm})",
                    ok,
                    f"Somme={fmt(soma)} | Total={fmt(t)} | diff={fmt(abs(soma-float(t)))}",
                    "error", "tab 3.1 — contrôle type"
                ))

    # ── CHECK C: contrôle identification (row 22)
    # D22=D18=D15+D16+D17
    # 3.1.1: Total général pros (Annonces) = Pros identifiés + Pros à identifier + Annonces incomplètes
    r_pros_id   = _311_row(ws311, 'Pros identifiés Joreca')
    r_pros_nid  = _311_row(ws311, 'Pros à identifier Joreca')
    r_incomplet = _311_row(ws311, 'Annonces incomplètes')

    if all(r is not None for r in [r_total_gen, r_pros_id, r_pros_nid, r_incomplet]):
        for site in SITES:
            if site not in site_cols: continue
            ann_col = site_cols[site][1]
            tg = ws311.cell(r_total_gen,  ann_col).value
            pi = ws311.cell(r_pros_id,    ann_col).value
            pn = ws311.cell(r_pros_nid,   ann_col).value
            ai = ws311.cell(r_incomplet,  ann_col).value
            vals = [tg, pi, pn, ai]
            if all(isinstance(v, (int,float)) for v in vals) and float(tg) > 0:
                soma = float(pi) + float(pn) + float(ai)
                ok   = (soma == float(tg))
                checks.append(check(
                    f"[tab 1 row 22] {site}: Pros id+non id+incomplets = Total général ({lm})",
                    ok,
                    f"Somme={fmt(soma)} | Total={fmt(tg)} | diff={fmt(abs(soma-float(tg)))}",
                    "error", "tab 3.1 — contrôle identification"
                ))

    # ── CHECK D: contrôle type sur 3.1.4 (segments Agences/Interm/Notaires/Autres)
    # tab 3.1 row 21: annonceurs totaux = somme des sous-types
    lc314, _ = get_lc(ws314)
    if lc314:
        total_all = row_val(ws314, 'Total', lc314, skip=0)
        agences   = row_val(ws314, 'Total', lc314, skip=1)
        interm314 = row_val(ws314, 'Total', lc314, skip=2)
        notaires  = row_val(ws314, 'Total', lc314, skip=3)
        autres314 = row_val(ws314, 'Total', lc314, skip=4)
        if all(v is not None for v in [total_all, agences, interm314, notaires, autres314]):
            soma = agences + interm314 + notaires + autres314
            checks.append(check(
                f"[tab 1 row 21] 3.1.4 Agences+Interméd+Notaires+Autres = Total annonceurs ({lm})",
                soma == total_all,
                f"Somme={fmt(soma)} | Total={fmt(total_all)} | diff={fmt(abs(soma-total_all))}",
                "error", "tab 3.1 — contrôle type (3.1.4)"
            ))

    return checks

# ─── TAB 3.2 CHECKS ──────────────────────────────────────────────────────────
# Source des formules Panel Checker (vérifiées):
#
# tab 3.2 row 24 "contrôle tab 3.1.4":
#   K24: K22='tab 3.1'!$O194  (Leboncoin)
#   AA24: AA22='tab 3.1'!$O202  (Total)
#   → 3.2.1 TOTAL row per site == 3.1.4 "Annonceurs professionnels identifiés"[site, lc]
#     (tab 3.1 O190-O202 = 3.1.4 section 1 rows 6-20)
#
# tab 3.2 row 25 "contrôle segment":
#   K25: K22=K45+K69+K91+K113
#   → 3.2.1 grand TOTAL = Agences TOTAL + Intermédiaires TOTAL + Notaires TOTAL + Autres TOTAL
#
# tab 3.2 row 47 "contrôle tab 3.1.4" (Agences):
#   K47: K45='tab 3.1'!$O216
#   → 3.2.1 Agences TOTAL per site == 3.1.4 Agences[site, lc]
#
# tab 3.2 row 48 "contrôle tab 3.2.2":
#   K48: K45=K348
#   → 3.2.1 Agences TOTAL == 3.2.2 Agences total per département (sum)

def _321_site_cols(ws321):
    """3.2.1: retourne dict site→Pros_col (row 6 = sites, row 7 = Pros/Poids)"""
    cols = {}
    for c in range(3, ws321.max_column+1):
        raw = ws321.cell(6, c).value
        if raw and isinstance(raw, str):
            name = raw.strip().strip("'\"")
            h7 = ws321.cell(7, c).value
            if h7 and str(h7).strip() == 'Pros':
                cols[name] = c
    return cols

def _321_total_rows(ws321):
    """3.2.1: retourne liste des row indices avec label 'TOTAL' col 2"""
    rows = []
    for r in range(1, ws321.max_row+1):
        v = ws321.cell(r, 2).value
        if v and str(v).strip().upper() == 'TOTAL':
            rows.append(r)
    return rows

def check_tab32(wb32_bytes, wb31_bytes):
    checks = []
    wb32 = load_workbook(io.BytesIO(wb32_bytes), data_only=True)
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)

    ws321 = wb32['3.2.1 Pros par régions']
    ws314 = wb31['3.1.4 Evolution Pros par type']
    lc314, lm = get_lc(ws314)
    if not lc314: return checks

    site_cols = _321_site_cols(ws321)
    total_rows = _321_total_rows(ws321)

    SITES = ['AvendreAlouer', "Bien'ici", 'Figaro Immo', 'GreenAcres', 'Leboncoin',
             'LogicImmo', 'MeilleursAgents', 'OuestFrance', 'PAP', 'ParuVendu',
             'SeLoger', 'Superimmo', 'Total']
    ALIAS_314 = {'GreenAcres': 'Green-Acres'}

    # ── CHECK A: contrôle tab 3.1.4 (row 24) ──
    # 3.2.1 TOTAL (section 1, grand total row) per site == 3.1.4 site row, lc314
    # tab 3.1 O190-O202 = 3.1.4 section 1 "Annonceurs professionnels identifiés":
    #   row 190=AvendreAlouer, 191=Bien'ici, ..., 202=Total, 204=Total Dédupliqué
    if total_rows:
        grand_row = total_rows[0]
        for site in SITES:
            site_314 = ALIAS_314.get(site, site)
            v314 = row_val(ws314, site_314, lc314)
            v321 = site_cols.get(site) and ws321.cell(grand_row, site_cols[site]).value

            if v314 is not None and v321 is not None and float(v314) > 0:
                ok = (float(v314) == float(v321))
                checks.append(check(
                    f"[tab 3.2 row 24] {site}: 3.2.1 TOTAL = 3.1.4 annonceurs ({lm})",
                    ok,
                    f"3.1.4={fmt(v314)} | 3.2.1={fmt(v321)} | diff={fmt(abs(float(v314)-float(v321)))}",
                    "error", "tab 3.2 — contrôle tab 3.1.4"
                ))

    # ── CHECK B: contrôle segment (row 25) ──
    # K25: K22=K45+K69+K91+K113
    # total_rows[0]=all pros, [1]=Agences, [2]=Intermédiaires, [3]=Notaires, [4]=Autres
    if len(total_rows) >= 5:
        for site in SITES:
            col = site_cols.get(site)
            if col is None: continue
            gt  = ws321.cell(total_rows[0], col).value
            ag  = ws321.cell(total_rows[1], col).value
            it  = ws321.cell(total_rows[2], col).value
            nt  = ws321.cell(total_rows[3], col).value
            au  = ws321.cell(total_rows[4], col).value
            vals = [gt, ag, it, nt, au]
            if all(isinstance(v, (int,float)) for v in vals) and float(gt) > 0:
                soma = float(ag)+float(it)+float(nt)+float(au)
                ok = (soma == float(gt))
                checks.append(check(
                    f"[tab 3.2 row 25] {site}: 3.2.1 Grand Total = Agences+Interméd+Notaires+Autres ({lm})",
                    ok,
                    f"Somme={fmt(soma)} | Grand Total={fmt(gt)} | diff={fmt(abs(soma-float(gt)))}",
                    "error", "tab 3.2 — contrôle segment"
                ))

    # ── CHECK C: contrôle tab 3.1.4 par section (row 47) ──
    # K47: K45='tab 3.1'!$O216 (Agences Bien'ici)
    # 3.1.4 section Agences (rows 28-43): row 40=Total, rows 28-39=sites
    SECTIONS_3_1_4 = [
        ('Agences immobilières',  1, 'tab 3.2 — contrôle tab 3.1.4 (Agences)'),
        ('Intermédiaires',        2, 'tab 3.2 — contrôle tab 3.1.4 (Interméd.)'),
        ('Notaires',              3, 'tab 3.2 — contrôle tab 3.1.4 (Notaires)'),
        ('Autres annonceurs',     4, 'tab 3.2 — contrôle tab 3.1.4 (Autres)'),
    ]
    if len(total_rows) >= 5:
        for sec_label, sec_idx, sheet_label in SECTIONS_3_1_4:
            sec_row_321 = total_rows[sec_idx]
            for site in SITES:
                site_314 = ALIAS_314.get(site, site)
                # 3.1.4 section rows: skip=sec_idx finds the section's site row
                v314 = row_val(ws314, site_314, lc314, skip=sec_idx)
                col = site_cols.get(site)
                if col is None or v314 is None: continue
                v321 = ws321.cell(sec_row_321, col).value
                if v321 is not None and float(v314) > 0:
                    ok = (float(v314) == float(v321))
                    checks.append(check(
                        f"[tab 3.2 row 47+] {sec_label} {site}: 3.2.1 = 3.1.4 ({lm})",
                        ok,
                        f"3.1.4={fmt(v314)} | 3.2.1={fmt(v321)} | diff={fmt(abs(float(v314)-float(v321)))}",
                        "error", sheet_label
                    ))

    # ── CHECK D: contrôle MAX <= Total Dédupliqué (tab 3.2 row 68/473)
    # Formule: =MAX(C68,...,Y68)<=AG68
    # Pour chaque région: max(sites) doit être <= Total Dédupliqué
    # Si MAX > Dédup → erreur de déduplication dans les données sources
    dedup_col_321 = None
    for c in range(3, ws321.max_column+1):
        h6 = ws321.cell(6, c).value
        h7 = ws321.cell(7, c).value
        if h6 and 'Dédupliqué' in str(h6) and h7 and 'Pros' in str(h7):
            dedup_col_321 = c; break
    
    site_cols_list = sorted(site_cols.values())
    if dedup_col_321 and site_cols_list:
        violations = []
        for r in range(8, ws321.max_row+1):
            region = ws321.cell(r, 2).value
            if not region or not isinstance(region, str): continue
            if str(region).strip().upper() in ('TOTAL', 'SITE', 'DÉPARTEMENT', 'RÉGION', ''): continue
            
            site_vals = [ws321.cell(r, c).value for c in site_cols_list
                        if isinstance(ws321.cell(r, c).value, (int,float)) and ws321.cell(r, c).value > 0]
            dedup_val = ws321.cell(r, dedup_col_321).value
            
            if site_vals and isinstance(dedup_val, (int,float)) and dedup_val >= 0:
                max_val = max(site_vals)
                if max_val > dedup_val:
                    violations.append(f"{region.strip()} (MAX={fmt(max_val)} > Dédup={fmt(dedup_val)})")
        
        if violations:
            checks.append(check(
                f"[tab 3.2] MAX(sites) ≤ Total Dédupliqué pour chaque région ({lm})",
                False,
                f"❌ {len(violations)} région(s) avec incohérence : {', '.join(violations[:5])}"
                + (f" et {len(violations)-5} autres" if len(violations) > 5 else "")
                + " → Vérifier la déduplication dans le fichier 3.2",
                "error", "tab 3.2 — contrôle déduplication"
            ))
        else:
            checks.append(check(
                f"[tab 3.2] MAX(sites) ≤ Total Dédupliqué pour chaque région ({lm})",
                True, "✅ Toutes les régions sont cohérentes.", "error",
                "tab 3.2 — contrôle déduplication"
            ))

    return checks

def check_tab5_vs_tab32(wb5_bytes, wb32_bytes):
    """
    Panel Checker tab 5 satır 347/360/386 — IDF/Alpes Maritimes pros:
    Formül: =H342='tab 3.2'!$AE$256  →  tab5 Agences06 pros = 3.2.1 Alpes-Maritimes MeilleursAgents pros
    Formül: =H381='tab 3.2'!$AE$634  →  tab5 Autres06 pros = 3.2.1 Autres Alpes-Maritimes pros
    """
    checks = []
    try:
        wb5  = load_workbook(io.BytesIO(wb5_bytes), data_only=True)
        wb32 = load_workbook(io.BytesIO(wb32_bytes), data_only=True)
    except Exception:
        return checks

    ws5  = wb5[wb5.sheetnames[0]]
    if '3.2.2 Pros par département' not in wb32.sheetnames:
        return checks
    ws322 = wb32['3.2.2 Pros par département']

    # tab 5 (Focus IDF & Alpes-Maritimes) yapısı
    # Bulmamız gereken: Alpes-Maritimes (06) satırları per section
    # Panel Checker tab 5 H/I kolonları = bazı site değerleri
    # Karşılaştırma 3.2.2'deki Alpes-Maritimes departmanıyla yapılıyor

    # 3.2.2'de Alpes-Maritimes (06) satırını bul
    alpes_row_322 = None
    for r in range(1, min(ws322.max_row+1, 200)):
        v = ws322.cell(r, 1).value
        if v and ('alpes' in str(v).lower() or '06' in str(v)):
            alpes_row_322 = r; break

    if not alpes_row_322:
        return checks

    lm = None
    # 3.2.2 header satırından son ay ve site kolonlarını bul
    for r in range(1, 10):
        for c in range(1, ws322.max_column+1):
            v = ws322.cell(r, c).value
            if isinstance(v, datetime.datetime):
                lm = v.strftime('%b-%y'); break
        if lm: break

    # tab 5 kaynak dosyasında Alpes-Maritimes satırlarını bul
    alpes_rows_5 = []
    for r in range(1, ws5.max_row+1):
        v = ws5.cell(r, 2).value or ws5.cell(r, 1).value
        if v and ('alpes' in str(v).lower() or '06' in str(v) or 'maritimes' in str(v).lower()):
            alpes_rows_5.append((r, str(v)))

    # tab 5 son ay col
    lc5, lm5 = get_lc(ws5)
    # tab 3.2.2 son ay col
    lc322, _ = get_lc(ws322)

    if lc5 and lc322 and alpes_rows_5:
        for r5, label5 in alpes_rows_5[:5]:
            v5 = ws5.cell(r5, lc5).value
            v322 = ws322.cell(alpes_row_322, lc322).value
            if isinstance(v5, (int,float)) and isinstance(v322, (int,float)) and v5 > 0:
                ok = (v5 == v322)
                checks.append(check(
                    f"[tab 5 Alpes-Maritimes] {label5}: fichier 5 = 3.2.2 Alpes-Maritimes ({lm5 or lm})",
                    ok,
                    (f"✅ Cohérent : {fmt(v5)}" if ok else
                     f"❌ tab5={fmt(v5)} | 3.2.2 Alpes-Maritimes={fmt(v322)} | diff={fmt(abs(v5-v322))}"
                     f" → Vérifier les fichiers 5 et 3.2"),
                    "error", "tab 5 — contrôle vs tab 3.2"
                ))

    return checks

def check_tab411_vs_tab1(wb41_bytes, wb1_bytes):
    """
    tab 4.1.1 rows 171/193/215/237 — 'contrôle tab 1-4' (hardcoded False dans Panel Checker)
    Formül olmasi gereken: 4.1.1 Pros Ventes + Pros Locations per site = 1.1 Ancien Pros per site
    Section yapısı:
      4.1.1 row 80 TOTAL = 'Ancien - Total Annonces de Ventes de Professionnels'
      4.1.1 row 100 TOTAL = 'Ancien - Total Annonces de Locations de Professionnels'
      1.1 rows 27-42 = Annonces Ancien Pros per site (dernier col)
    """
    checks = []
    try:
        wb41 = load_workbook(io.BytesIO(wb41_bytes), data_only=True)
        wb1  = load_workbook(io.BytesIO(wb1_bytes), data_only=True)
    except Exception:
        return checks

    if '4.1.1 Régions - Annonces' not in wb41.sheetnames: return checks
    ws411 = wb41['4.1.1 Régions - Annonces']
    ws11  = wb1['1.1 Total']
    lc1, lm = get_lc(ws11)
    if not lc1: return checks

    SITES = ['AvendreAlouer', "Bien'ici", 'Figaro Immo', 'Green-Acres', 'Leboncoin',
             'LogicImmo', 'MeilleursAgents', 'OuestFrance', 'PAP', 'ParuVendu',
             'SeLoger', 'Superimmo']
    ALIAS_411 = {'GreenAcres': 'Green-Acres', 'SuperImmo': 'Superimmo'}
    ALIAS_11  = {'SuperImmo': 'Superimmo'}

    # 4.1.1 site column mapping (row 5 = headers)
    site_col_411 = {}
    for c in range(3, ws411.max_column+1):
        h = ws411.cell(5, c).value
        if h:
            name = ALIAS_411.get(str(h).strip(), str(h).strip())
            site_col_411[name] = c

    # 4.1.1 TOTAL row per section
    # row 80 = Pros Ventes TOTAL, row 100 = Pros Locations TOTAL
    # Verification: find exact TOTAL rows by scanning
    section_total_rows = {}
    section_headers = {
        80:  'Pros Ventes',
        100: 'Pros Locations',
        120: 'Particuliers Ventes',
        140: 'Particuliers Locations',
    }
    # Verify these are indeed TOTAL rows
    for expected_row, label in section_headers.items():
        b = ws411.cell(expected_row, 2).value
        if b and str(b).strip().upper() == 'TOTAL':
            section_total_rows[label] = expected_row

    pros_vte_row = section_total_rows.get('Pros Ventes')
    pros_loc_row = section_total_rows.get('Pros Locations')
    if not pros_vte_row or not pros_loc_row: return checks

    # 1.1 site row mapping in Annonces Ancien section (rows ~27-44)
    site_row_11 = {}
    for r in range(24, 50):
        b = ws11.cell(r, 2).value
        if b and isinstance(b, str):
            name = ALIAS_11.get(b.strip(), b.strip())
            if name in SITES:
                site_row_11[name] = r

    # Per site: 4.1.1 Pros Ventes + Locations = 1.1 Ancien Pros
    errors = []
    ok_count = 0
    for site in SITES:
        col_411 = site_col_411.get(site)
        row_11  = site_row_11.get(site)
        if col_411 is None or row_11 is None: continue

        v_vte = ws411.cell(pros_vte_row, col_411).value or 0
        v_loc = ws411.cell(pros_loc_row, col_411).value or 0
        v_sum = float(v_vte) + float(v_loc)
        v_11  = float(ws11.cell(row_11, lc1).value or 0)

        if v_sum == v_11:
            ok_count += 1
        else:
            errors.append(
                f"{site}: 4.1.1 Ventes({fmt(v_vte)})+Loc({fmt(v_loc)})={fmt(v_sum)} ≠ 1.1={fmt(v_11)} (diff={fmt(abs(v_sum-v_11))})"
            )

    if errors:
        checks.append(check(
            f"[tab 4.1.1 contrôle tab 1] 4.1.1 Pros Ventes+Locations = 1.1 Ancien Pros par site ({lm})",
            False,
            "❌ " + " | ".join(errors[:4]) + (f" +{len(errors)-4} autres" if len(errors) > 4 else "")
            + " → Vérifier les fichiers 4.1 et 1",
            "error", "tab 4.1.1 — contrôle tab 1-4 (rows 171/193/215/237)"
        ))
    elif ok_count > 0:
        checks.append(check(
            f"[tab 4.1.1 contrôle tab 1] 4.1.1 Pros Ventes+Locations = 1.1 Ancien Pros par site ({lm})",
            True,
            f"✅ {ok_count} sites cohérents.",
            "error", "tab 4.1.1 — contrôle tab 1-4 (rows 171/193/215/237)"
        ))

    # Total Dedup check (AO kontrolü)
    td_vte = ws411.cell(pros_vte_row, 18).value or 0
    td_loc = ws411.cell(pros_loc_row, 18).value or 0
    td_sum = float(td_vte) + float(td_loc)
    # 1.1 row 42 = 'Total Panel Dédupliqué Marché' Ancien Pros
    td_11 = None
    for r in range(38, 50):
        b = ws11.cell(r, 2).value
        if b and 'dédupliqué marché' in str(b).lower():
            td_11 = ws11.cell(r, lc1).value; break

    if td_11 is not None:
        ok_td = (td_sum == float(td_11))
        checks.append(check(
            f"[tab 4.1.1 AO contrôle] Pros Ventes+Locations Total Dedup = 1.1 Total Dédupliqué Marché ({lm})",
            ok_td,
            (f"✅ Cohérent : {fmt(td_sum)}" if ok_td else
             f"❌ 4.1.1 Pros Vtes({fmt(td_vte)})+Loc({fmt(td_loc)})={fmt(td_sum)} ≠ 1.1={fmt(td_11)} (diff={fmt(abs(td_sum-float(td_11)))})"),
            "error", "tab 4.1.1 — contrôle tab 1-4 (rows 171/193/215/237)"
        ))

    return checks
    """
    tab 5-2 (Grand Ouest) — MAX(sites) <= Total Dédupliqué per département
    Panel Checker: =MAX(B460,...,F460)<=G460 — Vendée/Finistère/Maine-et-Loire hataları
    """
    checks = []
    try:
        wb52 = load_workbook(io.BytesIO(wb52_bytes), data_only=True)
    except Exception:
        return checks

    for sn in wb52.sheetnames:
        if sn.lower() == 'intro': continue
        ws = wb52[sn]
        lc, lm = get_lc(ws)
        if not lc: continue

        # Kolonlar: B=site1, C=site2,... dernier col avant Dédup = sites, son col = Dédup
        # Header row 2'de "Total" ya da "Dédupliqué" içeren kolon = dedup_col
        dedup_col = None
        site_cols = []
        for c in range(2, min(ws.max_column+1, 15)):
            h = ws.cell(2, c).value or ws.cell(1, c).value
            if not h: continue
            if isinstance(h, str) and ('dédup' in h.lower() or 'total' in h.lower()):
                dedup_col = c; break
            site_cols.append(c)

        if not site_cols or not dedup_col: continue

        violations = []
        for r in range(3, ws.max_row+1):
            label = ws.cell(r, 1).value
            if not label or not isinstance(label, str): continue
            if any(x in str(label).upper() for x in ['TOTAL', 'SITE']): continue

            site_vals = [ws.cell(r, c).value for c in site_cols
                        if isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value > 0]
            dedup_val = ws.cell(r, dedup_col).value

            if site_vals and isinstance(dedup_val, (int, float)) and dedup_val > 0:
                max_val = max(site_vals)
                if max_val > dedup_val:
                    violations.append(f"{str(label).strip()} (MAX={fmt(max_val)} > Dédup={fmt(dedup_val)})")

        if violations:
            checks.append(check(
                f"[tab 5-2] MAX(sites) ≤ Total Dédupliqué — {sn} ({lm})",
                False,
                f"❌ {len(violations)} département(s) : {', '.join(violations[:5])}"
                + (f" +{len(violations)-5}" if len(violations) > 5 else "")
                + " → Vérifier la déduplication dans le fichier 5.2",
                "error", "tab 5-2 — contrôle MAX déduplication"
            ))
        elif lm and site_cols:
            checks.append(check(
                f"[tab 5-2] MAX(sites) ≤ Total Dédupliqué — {sn} ({lm})",
                True, "✅ Tous les départements sont cohérents.", "error",
                "tab 5-2 — contrôle MAX déduplication"
            ))

    return checks


def check_tab5_max(wb5_bytes):
    """
    tab 5 (Focus IDF & Alpes-Maritimes) — MAX(sites) <= Total Dédupliqué per arrondissement
    Panel Checker: =MAX(C590,D590,E590,F590)<=G590 — 5ème et 13ème arr. en erreur
    """
    checks = []
    try:
        wb5 = load_workbook(io.BytesIO(wb5_bytes), data_only=True)
    except Exception:
        return checks

    for sn in wb5.sheetnames:
        if sn.lower() == 'intro': continue
        ws = wb5[sn]
        lc, lm = get_lc(ws)
        if not lc: continue

        dedup_col = None
        site_cols = []
        for c in range(2, min(ws.max_column+1, 12)):
            h = ws.cell(2, c).value or ws.cell(1, c).value
            if not h: continue
            if isinstance(h, str) and ('dédup' in h.lower() or 'dedup' in h.lower()):
                dedup_col = c; break
            if isinstance(h, str) and h.strip() not in ('', 'Site', 'Arrondissement'):
                site_cols.append(c)

        if not site_cols or not dedup_col: continue

        violations = []
        for r in range(3, ws.max_row+1):
            label = ws.cell(r, 1).value or ws.cell(r, 2).value
            if not label or not isinstance(label, str): continue
            label_str = str(label).strip()
            if any(x in label_str.upper() for x in ['TOTAL', 'SITE', 'RÉGION']): continue
            if not any(c.isdigit() for c in label_str): continue

            site_vals = [ws.cell(r, c).value for c in site_cols
                        if isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value > 0]
            dedup_val = ws.cell(r, dedup_col).value

            if site_vals and isinstance(dedup_val, (int, float)) and dedup_val > 0:
                max_val = max(site_vals)
                if max_val > dedup_val:
                    violations.append(f"{label_str} (MAX={fmt(max_val)} > Dédup={fmt(dedup_val)})")

        if violations:
            checks.append(check(
                f"[tab 5] MAX(sites) ≤ Total Dédupliqué — {sn} ({lm})",
                False,
                f"❌ {len(violations)} arrondissement(s) : {', '.join(violations[:5])}"
                + " → Vérifier la déduplication dans le fichier 5",
                "error", "tab 5 — contrôle MAX déduplication"
            ))
        elif lm and site_cols:
            checks.append(check(
                f"[tab 5] MAX(sites) ≤ Total Dédupliqué — {sn} ({lm})",
                True, "✅ Tous les arrondissements sont cohérents.", "error",
                "tab 5 — contrôle MAX déduplication"
            ))

    return checks


# ─── MAIN RUNNER ─────────────────────────────────────────────────────────────

def classify_files(uploaded_files):
    file_labels = {
        'file1':'1 — Evolution panel','file2':'2 — Performance qualité',
        'file3_1':'3.1 — Analyse Pros','file3_2':'3.2 — Géographique Pros',
        'file4_1':'4.1 — Stats géographiques','file4_2':'4.2 — Exclusivité/Partage',
        'file5':'5 — Focus IDF','file5_2':'5.2 — Grand Ouest',
    }
    files_bytes = {}
    for f in uploaded_files:
        nl = f.name.lower().replace(' ','_').replace('é','e').replace('è','e').replace('ô','o')
        fb = f.read()
        if '5_2' in nl or ('5' in nl and ('grand' in nl or 'ouest' in nl)):
            key = 'file5_2'
        elif '5' in nl and ('idf' in nl or 'ile' in nl or 'alpes' in nl or 'focus' in nl) and '5_2' not in nl:
            key = 'file5'
        elif '4_2' in nl or ('4' in nl and ('exclus' in nl or 'partag' in nl)):
            key = 'file4_2'
        elif '4_1' in nl or ('4' in nl and 'stat' in nl and '4_2' not in nl):
            key = 'file4_1'
        elif '3_2' in nl or ('3' in nl and 'geo' in nl):
            key = 'file3_2'
        elif '3_1' in nl or ('3' in nl and 'pros' in nl and '3_2' not in nl):
            key = 'file3_1'
        elif '2' in nl and ('perform' in nl or 'qualit' in nl):
            key = 'file2'
        elif '1' in nl and ('evolution' in nl or 'annonce' in nl) and '3_1' not in nl and '4_1' not in nl:
            key = 'file1'
        else:
            continue
        files_bytes[key] = fb
    return files_bytes, file_labels

def run_all_qc_gold(uploaded_files):
    files_bytes, file_labels = classify_files(uploaded_files)
    all_checks = []

    # 3.1 × 1 — contrôle type + contrôle tab 1
    if 'file3_1' in files_bytes and 'file1' in files_bytes:
        all_checks += check_tab31(files_bytes['file3_1'], files_bytes['file1'])
    # 3.2 × 3.1 — contrôle tab 3.1.4 + contrôle segment + MAX dédup
    if 'file3_2' in files_bytes and 'file3_1' in files_bytes:
        all_checks += check_tab32(files_bytes['file3_2'], files_bytes['file3_1'])
    # tab 5 × tab 3.2 — Alpes-Maritimes cross-check
    if 'file5' in files_bytes and 'file3_2' in files_bytes:
        all_checks += check_tab5_vs_tab32(files_bytes['file5'], files_bytes['file3_2'])
    # tab 4.1.1 × tab 1 — contrôle tab 1-4 (rows 171/193/215/237, hardcoded False dans Panel Checker)
    if 'file4_1' in files_bytes and 'file1' in files_bytes:
        all_checks += check_tab411_vs_tab1(files_bytes['file4_1'], files_bytes['file1'])
    if 'file5' in files_bytes:
        all_checks += check_tab5_max(files_bytes['file5'])
    # tab 5-2 — MAX(sites) <= Dédupliqué par département (Grand Ouest)
    if 'file5_2' in files_bytes:
        all_checks += check_tab52_max(files_bytes['file5_2'])

    return all_checks, files_bytes


def parse_ham_sections(uploaded_files):
    result = {}
    for f in uploaded_files:
        fb = f.read()
        wb = load_workbook(io.BytesIO(fb), data_only=True)
        for sn in wb.sheetnames:
            if sn == 'Intro': continue
            ws = wb[sn]
            for r in range(1, ws.max_row+1):
                b = ws.cell(r,2).value
                if b not in ('Site','Département'): continue
                month_cols, months = [], []
                for c in range(3, ws.max_column+1):
                    v = ws.cell(r,c).value
                    if v is None: continue
                    if (isinstance(v,datetime.datetime) or
                        (isinstance(v,(int,float)) and 40000<v<50000) or
                        (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
                        month_cols.append(c); months.append(excel_date_str(v))
                if len(months)<2: continue
                label = None
                for tr in range(r-1,max(0,r-5),-1):
                    v = ws.cell(tr,2).value
                    if v and isinstance(v,str) and len(v.strip())>3 and v.strip()!='Site':
                        label=v.strip(); break
                if not label: label=sn
                sites={}; totals={}
                for dr in range(r+1, min(r+35,ws.max_row+1)):
                    site=ws.cell(dr,2).value
                    if not site or not isinstance(site,str): break
                    site=site.strip()
                    if site=='Site': break
                    vals=[float(ws.cell(dr,c).value) if isinstance(ws.cell(dr,c).value,(int,float)) else None for c in month_cols]
                    real=[v for v in vals if v is not None and v>0]
                    if 'Total Panel' in site or site=='Total':
                        if len(real)>=2: totals[site]=vals
                    elif site not in SKIP_ROWS and site!='' and len(real)>=2:
                        sites[site]=vals
                if sites:
                    result[f"{sn}_{r}"] = {'sheet':sn,'label':label,'months':months,
                                           'sites':sites,'totals':totals,'file':f.name}
    return result

def analyze_trends(sections):
    rows = []
    for key, sec in sections.items():
        months = sec['months']
        total_dedup = {}
        for k,v in sec['totals'].items():
            if 'Dédupliqué' in k or 'Dedup' in k:
                total_dedup={i:x for i,x in enumerate(v) if x is not None and x>0}; break
        for site, vals in sec['sites'].items():
            real=[(i,v) for i,v in enumerate(vals) if v is not None and v>5]
            if len(real)<2: continue
            li,lv=real[-1]; pi,pv=real[-2]
            evol=(lv/pv-1)*100 if pv>0 else None
            pr_m=(lv/total_dedup[li]*100) if li in total_dedup and total_dedup[li]>0 else None
            pr_m1=(pv/total_dedup[pi]*100) if pi in total_dedup and total_dedup[pi]>0 else None
            evol_pr=(pr_m-pr_m1) if pr_m and pr_m1 else None
            status='ok'; flags=[]
            if evol is not None:
                if evol<=-20: status='critical'; flags.append(f"Monthly change: {evol:+.1f}% (vs prev month)")
                elif evol<=-10: status='warning'; flags.append(f"Monthly change: {evol:+.1f}% (vs prev month)")
                elif evol>=30: status='warning'; flags.append(f"Monthly surge: {evol:+.1f}%")
            max_hist=max((v for v in vals[:-1] if v and v>5),default=None)
            if max_hist and lv/max_hist<0.6:
                status='critical'; flags.append(f"Crash vs historical max: {(lv/max_hist-1)*100:+.1f}%")
            last3=[v for v in vals[-3:] if v and v>0]
            if len(last3)==3 and last3[0]>last3[1]>last3[2]:
                drop=(last3[2]-last3[0])/last3[0]*100
                if drop<-5:
                    if status=='ok': status='warning'
                    flags.append(f"3-month downtrend: {drop:+.1f}%")
            rows.append({'file':sec['file'],'sheet':sec['sheet'],'section':sec['label'],
                'site':site,'month_m':months[li] if li<len(months) else '?',
                'month_m1':months[pi] if pi<len(months) else '?',
                'val_m':lv,'val_m1':pv,'evol_pct':evol,'pr_m':pr_m,'evol_pr':evol_pr,
                'status':status,'flags':flags,'values':vals,'months':months})
    return rows

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## QC Monitor")
    st.divider()
    st.markdown("**Market**")
    sector = st.radio("", ["🏠 Real Estate","🚗 Auto"], label_visibility="collapsed", key="sector")
    if sector == "🏠 Real Estate":
        market = st.radio("", ["REFR"], label_visibility="collapsed", key="re_sub")
    else:
        st.caption("Coming soon"); market = None
    st.divider()
    st.markdown("**Page**")
    page = st.radio("", ["📊 Panel Tables Monitor","✅ Panel Checker (QC Gold)"],
        label_visibility="collapsed", key="page_sel")
    st.divider()
    st.markdown("**Upload files**")
    st.caption("Drop all files here — system sorts automatically")
    uploaded_files = st.file_uploader("", type=['xlsx'], accept_multiple_files=True,
        label_visibility="collapsed", key="global_upload")
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) loaded")

if market is None:
    st.markdown("# 🚗 Auto"); st.info("Coming soon"); st.stop()

# ─── PAGE: PANEL TABLES MONITOR ──────────────────────────────────────────────

if page == "📊 Panel Tables Monitor":
    st.markdown(f"# 📊 Panel Tables Monitor — {market}")
    st.divider()
    if not uploaded_files:
        st.info("👆 Upload your files in the sidebar to get started"); st.stop()

    with st.spinner("Loading..."):
        sections = parse_ham_sections(uploaded_files)
        rows = analyze_trends(sections)

    if not rows:
        st.warning("No data found."); st.stop()

    n_crit=sum(1 for r in rows if r['status']=='critical')
    n_warn=sum(1 for r in rows if r['status']=='warning')

    if n_crit>0: st.error(f"❌ REFUSED — {n_crit} critical · {n_warn} warnings")
    elif n_warn>3: st.warning(f"⚠️ TO REVIEW — {n_warn} warnings")
    else:
        st.success(f"✅ VALIDATED — {n_warn} minor warning(s)")
        if n_warn>0: st.info(f"{n_warn} warning(s) to monitor below")

    c1,c2 = st.columns([2,3])
    with c1:
        sev=st.radio("",["All","🔴 Critical","🟡 Warnings"],horizontal=True,key="sev_ptm")
    with c2:
        search=st.text_input("",placeholder="🔍 Search...",key="search_ptm",label_visibility="collapsed")

    filtered=rows
    if sev=="🔴 Critical": filtered=[r for r in rows if r['status']=='critical']
    elif sev=="🟡 Warnings": filtered=[r for r in rows if r['status'] in ('warning','critical')]
    if search:
        q=search.lower()
        filtered=[r for r in filtered if q in r['site'].lower() or q in r['section'].lower()]
    filtered=sorted(filtered,key=lambda r:{'critical':3,'warning':2,'ok':1}.get(r['status'],0),reverse=True)

    month_m=rows[0]['month_m'] if rows else ''
    month_m1=rows[0]['month_m1'] if rows else ''
    st.caption(f"M = {month_m}  ·  M-1 = {month_m1}  ·  {len(filtered)} series")

    table=[]
    for r in filtered:
        icon="🔴" if r['status']=='critical' else "🟡" if r['status']=='warning' else "✅"
        evol=f"{r['evol_pct']:+.1f}%" if r['evol_pct'] is not None else "—"
        pr=f"{r['pr_m']:.1f}%" if r['pr_m'] is not None else "—"
        evol_pr=f"{r['evol_pr']:+.2f}pp" if r['evol_pr'] is not None else "—"
        table.append({'':icon,'Section':r['section'][:35],'Site':r['site'],
            f"M ({month_m})":f"{r['val_m']:,.0f}",f"M-1 ({month_m1})":f"{r['val_m1']:,.0f}",
            'Evol %':evol,'Market share':pr,'MS evol':evol_pr,
            'Note':' | '.join(r['flags']) if r['flags'] else '—'})

    df=pd.DataFrame(table)
    def color_rows(row):
        if '🔴' in str(row.iloc[0]): return ['background-color:rgba(255,68,68,0.08)']*len(row)
        if '🟡' in str(row.iloc[0]): return ['background-color:rgba(245,166,35,0.05)']*len(row)
        return ['']*len(row)
    st.dataframe(df.style.apply(color_rows,axis=1),use_container_width=True,hide_index=True,
        height=min(600,40+35*len(table)))

    issue_rows=[r for r in filtered if r['flags']]
    if issue_rows:
        st.markdown("### Sites with anomalies")
        cols_n=3
        for i in range(0,min(len(issue_rows),12),cols_n):
            cols=st.columns(cols_n)
            for j,row in enumerate(issue_rows[i:i+cols_n]):
                with cols[j]:
                    clr='#ff4444' if row['status']=='critical' else '#f5a623'
                    evol_s=f"{row['evol_pct']:+.1f}%" if row['evol_pct'] else ''
                    icon2='🔴' if row['status']=='critical' else '🟡'
                    st.markdown(f"**{icon2} {row['site']}** — <span style='color:{clr}'>{evol_s}</span>",unsafe_allow_html=True)
                    st.caption(row['section'][:40])
                    vals=[v if v else 0 for v in row['values']]
                    fig=go.Figure(go.Scatter(x=row['months'][:len(vals)],y=vals,mode='lines+markers',
                        line=dict(color=clr,width=2),
                        marker=dict(size=[6 if k==len(vals)-1 else 0 for k in range(len(vals))]),
                        connectgaps=True,hovertemplate='%{x}: %{y:,.0f}<extra></extra>'))
                    fig.update_layout(height=120,margin=dict(l=0,r=0,t=0,b=0),
                        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False,tickfont=dict(size=8),nticks=4),
                        yaxis=dict(showgrid=True,gridcolor='rgba(128,128,128,0.15)',
                            tickfont=dict(size=8),tickformat='.2s'),showlegend=False)
                    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False},
                        key=f"ptm_{i}_{j}_{row['site'][:8]}")
                    for flag in row['flags']:
                        st.caption(f"⚠️ {flag}")

    st.divider()
    export=[{'File':r['file'],'Sheet':r['sheet'],'Section':r['section'],'Site':r['site'],
        'Status':r['status'],'Month M':r['month_m'],'Value M':r['val_m'],
        'Month M-1':r['month_m1'],'Value M-1':r['val_m1'],
        'Evol %':f"{r['evol_pct']:+.1f}%" if r['evol_pct'] else '',
        'Market share':f"{r['pr_m']:.1f}%" if r['pr_m'] else '',
        'Flags':' | '.join(r['flags'])} for r in rows]
    csv=pd.DataFrame(export).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download report (CSV)",data=csv,
        file_name=f"ptm_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",mime='text/csv')

# ─── PAGE: PANEL CHECKER (QC GOLD) ───────────────────────────────────────────

elif page == "✅ Panel Checker (QC Gold)":
    st.markdown(f"# ✅ Panel Checker (QC Gold) — {market}")
    st.caption("Upload the 10 source files — same checks as the Panel Checker, no Panel Checker file needed.")
    st.divider()

    if not uploaded_files:
        st.info("👆 Upload your source files in the sidebar to get started"); st.stop()

    with st.spinner("Running QC Gold checks..."):
        all_checks, files_bytes = run_all_qc_gold(uploaded_files)

    # ── Fichiers détectés ──
    FILE_LABEL_MAP = {
        'file1':'1 — Evolution panel','file2':'2 — Performance qualité',
        'file3_1':'3.1 — Analyse Pros','file3_2':'3.2 — Géographique Pros',
        'file4_1':'4.1 — Stats géographiques','file4_2':'4.2 — Exclusivité/Partage',
        'file5':'5 — Focus IDF','file5_2':'5.2 — Grand Ouest',
    }
    EXPECTED = list(FILE_LABEL_MAP.keys())
    detected_keys = list(files_bytes.keys())
    missing = [FILE_LABEL_MAP[k] for k in EXPECTED if k not in detected_keys]

    if missing:
        st.warning(f"⚠️ {len(missing)} file(s) not detected: {', '.join(missing)}")
    else:
        st.success(f"✅ All {len(detected_keys)} source files detected")

    with st.expander("📂 Detected files", expanded=False):
        for k in EXPECTED:
            icon = "✅" if k in detected_keys else "❌"
            st.markdown(f"{icon} {FILE_LABEL_MAP[k]}")

    if not all_checks:
        st.warning("No checks could be run. Make sure the correct files are uploaded."); st.stop()

    st.divider()

    errors   = [c for c in all_checks if c['status']=='fail' and c['severity']=='error']
    warnings = [c for c in all_checks if c['severity']=='warning']
    ok_checks= [c for c in all_checks if c['status']=='ok']

    # ── Global verdict ──
    if errors:
        st.error(f"❌ **{len(errors)} error(s)** — Files need correction before sending to the team")
    elif warnings:
        st.warning(f"⚠️ **{len(warnings)} warning(s)** — Review before validating")
    else:
        st.success(f"✅ **All {len(ok_checks)} checks passed** — Files are clean")

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total checks", len(all_checks))
    c2.metric("✅ Passed",    len(ok_checks))
    c3.metric("❌ Errors",    len(errors))
    c4.metric("🟡 Warnings",  len(warnings))

    st.divider()

    # ── Résultats par section (groupés par sheet) ──
    by_sheet = defaultdict(list)
    for c in all_checks:
        by_sheet[c.get('sheet','Other')].append(c)

    def sheet_priority(items):
        if any(i['status']=='fail' and i['severity']=='error' for i in items): return 0
        if any(i['severity']=='warning' for i in items): return 1
        return 2

    for sheet, items in sorted(by_sheet.items(), key=lambda x: sheet_priority(x[1])):
        n_err  = sum(1 for i in items if i['status']=='fail' and i['severity']=='error')
        n_warn = sum(1 for i in items if i['severity']=='warning')
        n_ok   = sum(1 for i in items if i['status']=='ok')
        icon   = "❌" if n_err else "🟡" if n_warn else "✅"
        badge  = f"{n_err} error(s)" if n_err else f"{n_warn} warning(s)" if n_warn else f"{n_ok} OK"

        with st.expander(f"{icon} **{sheet}** — {badge}", expanded=(n_err > 0)):
            for item in items:
                if item['status'] == 'ok':
                    st.markdown(f"✅ {item['check']}")
                elif item['severity'] == 'error':
                    st.markdown(f"❌ **{item['check']}**")
                    st.caption(f"   {item['detail']}")
                else:
                    st.markdown(f"🟡 {item['check']}")
                    st.caption(f"   {item['detail']}")

    st.divider()

    # ── Message taslağı (sadece hata varsa) ──
    if errors or warnings:
        st.markdown("### 📧 Message draft for the team")
        month_guess = ""
        for f in uploaded_files:
            n = f.name.lower()
            for m in ['janvier','février','mars','avril','mai','juin',
                      'juillet','août','septembre','octobre','novembre','décembre']:
                if m in n:
                    month_guess = m.capitalize() + " 2026"; break
            if month_guess: break
        if not month_guess: month_guess = "mois courant"

        lines = [f"Bonjour,\n",
                 f"En vérifiant les fichiers sources pour {month_guess}, "
                 f"j'ai détecté les points suivants à corriger :\n"]
        for item in errors:
            lines.append(f"❌ [{item.get('sheet','')}] {item['check']}\n   → {item['detail']}")
        for item in warnings:
            lines.append(f"⚠️ [{item.get('sheet','')}] {item['check']}\n   → {item['detail']}")
        lines.append("\nMerci de vérifier et corriger avant validation.\n\nCordialement")
        msg = "\n".join(lines)
        st.text_area("", value=msg, height=320, key="msg_qc_gold")
        st.download_button("⬇️ Download message (.txt)", data=msg.encode("utf-8"),
            file_name=f"qc_errors_{datetime.datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain")

    # ── CSV Export ──
    rows_exp = [{'Sheet': c.get('sheet',''), 'Check': c['check'],
        'Status': '❌ Error' if c['status']=='fail' and c['severity']=='error'
                  else '🟡 Warning' if c['severity']=='warning' else '✅ OK',
        'Detail': c.get('detail','')} for c in all_checks]
    csv = pd.DataFrame(rows_exp).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download full report (CSV)", data=csv,
        file_name=f"qc_gold_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime='text/csv')
